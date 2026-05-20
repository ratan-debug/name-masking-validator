from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd

try:
    import regex as re
except ModuleNotFoundError:  # pragma: no cover - fallback for minimal local runtimes
    import re  # type: ignore

try:
    from rapidfuzz import fuzz

    def fuzzy_ratio(left: str, right: str) -> float:
        return float(fuzz.ratio(left, right))

    def fuzzy_partial_ratio(left: str, right: str) -> float:
        return float(fuzz.partial_ratio(left, right))

except ModuleNotFoundError:  # pragma: no cover - app requirements install rapidfuzz
    from difflib import SequenceMatcher

    def fuzzy_ratio(left: str, right: str) -> float:
        return SequenceMatcher(None, left, right).ratio() * 100

    def fuzzy_partial_ratio(left: str, right: str) -> float:
        if not left or not right:
            return 0.0
        shorter, longer = (left, right) if len(left) <= len(right) else (right, left)
        best = 0.0
        for start in range(0, max(1, len(longer) - len(shorter) + 1)):
            window = longer[start : start + len(shorter)]
            best = max(best, fuzzy_ratio(shorter, window))
        return best


ORIGINAL_COLUMN = "Original Name"
MASKED_COLUMN = "Masked Name"
ORIGINAL_COLUMN_ALIASES = (
    ORIGINAL_COLUMN,
    "Input Owner Name",
    "Input Name",
)
MASKED_COLUMN_ALIASES = (
    MASKED_COLUMN,
    "Owner Name",
    "Masked Owner Name",
)
STATUS_COLUMN = "Match Status"
CONFIDENCE_COLUMN = "Confidence %"
MATCH_DETAILS_COLUMN = "Match Details"
FINAL_OUTPUT_COLUMN = "Final Output"

FULL_MATCH = "FULL_MATCH"
PARTIAL_MATCH = "PARTIAL_MATCH"
NO_MATCH = "NO_MATCH"


class ExcelProcessingError(ValueError):
    """Raised when the uploaded workbook cannot be processed."""


@dataclass(frozen=True)
class MatchResult:
    status: str
    confidence: int
    final_output: str
    details: str


@dataclass(frozen=True)
class SpanMatch:
    masked_index: int
    original_start: int
    original_end: int
    score: float
    output_token: str


@dataclass(frozen=True)
class Reconstruction:
    output: str
    matched_masked_count: int
    average_score: float


def process_excel(input_path: str | Path, output_path: str | Path) -> dict[str, int | float]:
    """Read an Excel file, append match analysis columns, and write a new workbook."""
    input_path = Path(input_path)
    output_path = Path(output_path)

    try:
        df = pd.read_excel(input_path, engine="openpyxl")
    except Exception as exc:  # pandas/openpyxl provide noisy exception types
        raise ExcelProcessingError(
            "Could not read the Excel file. Upload a valid .xlsx or .xlsm workbook."
        ) from exc

    if df.empty:
        raise ExcelProcessingError("The uploaded workbook has no data rows to process.")

    original_col, masked_col = find_name_columns(df.columns)

    results: list[MatchResult] = []
    for original_name, masked_name in df[[original_col, masked_col]].itertuples(index=False, name=None):
        results.append(analyze_name_pair(original_name, masked_name))

    df[STATUS_COLUMN] = [result.status for result in results]
    df[CONFIDENCE_COLUMN] = [result.confidence for result in results]
    df[MATCH_DETAILS_COLUMN] = [result.details for result in results]
    df[FINAL_OUTPUT_COLUMN] = [result.final_output for result in results]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Processed")
            autosize_columns(writer, df, "Processed")
    except Exception as exc:
        raise ExcelProcessingError("Could not write the processed Excel file.") from exc

    total = len(results)
    full = sum(1 for result in results if result.status == FULL_MATCH)
    partial = sum(1 for result in results if result.status == PARTIAL_MATCH)
    no_match = sum(1 for result in results if result.status == NO_MATCH)
    average_confidence = round(sum(result.confidence for result in results) / max(total, 1), 2)

    return {
        "total_rows": total,
        "full_matches": full,
        "partial_matches": partial,
        "no_matches": no_match,
        "average_confidence": average_confidence,
    }


def find_name_columns(columns: Iterable[object]) -> tuple[str, str]:
    names = [str(column) for column in columns]
    normalized = {normalize_column_name(name): name for name in names}

    original = first_matching_column(normalized, ORIGINAL_COLUMN_ALIASES)
    masked = first_matching_column(normalized, MASKED_COLUMN_ALIASES)
    if original and masked:
        return original, masked

    if len(names) >= 3:
        return names[1], names[2]

    if len(names) >= 2:
        return names[0], names[1]

    raise ExcelProcessingError(
        "The workbook must contain Original Name and Masked Name columns."
    )


def first_matching_column(normalized: dict[str, str], aliases: Iterable[str]) -> str | None:
    for alias in aliases:
        column = normalized.get(normalize_column_name(alias))
        if column:
            return column
    return None


def normalize_column_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def analyze_name_pair(original_value: object, masked_value: object) -> MatchResult:
    original = clean_cell(original_value)
    masked = clean_cell(masked_value)

    if not original and not masked:
        return MatchResult(NO_MATCH, 0, "", "empty original and masked values")
    if not original:
        return MatchResult(NO_MATCH, 0, masked, build_match_details([], tokenize_name(masked, keep_stars=True), [], "missing original"))
    if not masked:
        clean_original = " ".join(tokenize_name(original, keep_stars=False))
        return MatchResult(NO_MATCH, 45, clean_original, "missing masked value")

    original_tokens = tokenize_name(original, keep_stars=False)
    masked_tokens = tokenize_name(masked, keep_stars=True)

    if not original_tokens or not masked_tokens:
        confidence = int(round(fuzzy_ratio(normalize_name(original), normalize_name(masked))))
        status = FULL_MATCH if confidence >= 85 else NO_MATCH
        final_output = " ".join(original_tokens) if status == FULL_MATCH else masked
        return MatchResult(status, confidence, final_output, "could not tokenize one side")

    if "*" not in masked and literal_entity_match(original, masked):
        final_output = format_literal_output(masked)
        details = build_literal_match_details(original, masked, "confirmed readable alias/entity alignment")
        return MatchResult(FULL_MATCH, 100, final_output, details)

    matches = greedy_span_matches(masked_tokens, original_tokens)
    strong_matches = [match for match in matches if match.score >= 68]

    covered_original_indexes = {
        index
        for match in strong_matches
        for index in range(match.original_start, match.original_end)
    }
    original_coverage = len(covered_original_indexes) / max(len(original_tokens), 1)
    masked_coverage = len({match.masked_index for match in strong_matches}) / max(len(masked_tokens), 1)
    effective_masked_count = count_matchable_tokens(masked_tokens)
    masked_coverage = len({match.masked_index for match in strong_matches}) / max(effective_masked_count, 1)
    word_count_score = word_count_similarity(len(original_tokens), len(masked_tokens))
    initial_score = initials_similarity(original_tokens, masked_tokens)
    row_score = row_visible_similarity(original, masked)
    aligned_score = average(match.score for match in matches)

    confidence = clamp_int(
        aligned_score * 0.55
        + row_score * 0.20
        + initial_score * 0.15
        + word_count_score * 0.10
    )

    if (
        original_coverage >= 0.85
        and masked_coverage >= 0.75
        and word_count_score >= 65
        and confidence >= 72
    ):
        final_output = reconstruct_masked_order_output(masked_tokens, strong_matches)
        final_output = normalize_prefix_output(masked_tokens, final_output)
        details = build_match_details(original_tokens, masked_tokens, strong_matches, "confirmed full alignment")
        return MatchResult(FULL_MATCH, confidence, final_output, details)

    best_token_score = max((match.score for match in matches), default=0.0)
    flexible_reconstruction = reconstruct_masked_output(original_tokens, masked_tokens)
    flexible_coverage = flexible_reconstruction.matched_masked_count / max(effective_masked_count, 1)
    has_unmatched_specific_terms = count_unmatched_specific_terms(original_tokens, masked_tokens) >= 2

    if best_token_score < 48 and row_score < 38 and initial_score < 30 and flexible_reconstruction.matched_masked_count == 0:
        details = build_match_details(original_tokens, masked_tokens, matches, "no reliable token alignment")
        return MatchResult(NO_MATCH, clamp_int(max(best_token_score, row_score)), NO_MATCH, details)

    final_output = flexible_reconstruction.output or reconstruct_partial_output(original_tokens, masked_tokens, strong_matches, matches)
    candidate_confidence = clamp_int(max(confidence, best_token_score * 0.82, flexible_reconstruction.average_score))
    unmatched_masked_tokens = [
        masked_tokens[index]
        for index in range(len(masked_tokens))
        if index not in {match.masked_index for match in strong_matches}
    ]
    unmatched_tokens_are_initials = bool(unmatched_masked_tokens) and all(
        len(token.replace("*", "")) <= 1 for token in unmatched_masked_tokens
    )
    status = (
        PARTIAL_MATCH
        if (
            candidate_confidence >= 68
            and (
                masked_coverage >= 0.75
                or (flexible_coverage >= 0.50 and not has_unmatched_specific_terms)
                or unmatched_tokens_are_initials
            )
        )
        else NO_MATCH
    )
    result_confidence = candidate_confidence if status == PARTIAL_MATCH else confidence
    rule = "strong partial/subset alignment" if status == PARTIAL_MATCH else "weak candidate only; kept as no match"
    details = build_match_details(original_tokens, masked_tokens, strong_matches or matches[:2], rule)
    if status == NO_MATCH:
        final_output = NO_MATCH
    return MatchResult(status, result_confidence, final_output, details)


def clean_cell(value: object) -> str:
    if pd.isna(value):
        return ""
    value = str(value).strip()
    value = re.sub(r"\s+", " ", value)
    return value


def normalize_name(value: str, keep_stars: bool = False) -> str:
    pattern = r"[^A-Z0-9*/& ]+" if keep_stars else r"[^A-Z0-9 ]+"
    normalized = re.sub(pattern, " ", value.upper())
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def compact(value: str, keep_stars: bool = False) -> str:
    pattern = r"[^A-Z0-9*]+" if keep_stars else r"[^A-Z0-9]+"
    return re.sub(pattern, "", value.upper())


def tokenize_name(value: str, keep_stars: bool) -> list[str]:
    normalized = normalize_name(value, keep_stars=keep_stars)
    tokens = []
    for token in normalized.split():
        if keep_stars:
            if token == "&" or re.search(r"[A-Z0-9*]", token):
                tokens.append(token)
        elif re.search(r"[A-Z0-9]", token):
            tokens.append(token)
    return tokens


def literal_entity_match(original: str, masked: str) -> bool:
    original_tokens = canonical_entity_tokens(original)
    masked_tokens = canonical_entity_tokens(masked)
    if not original_tokens or not masked_tokens:
        return False
    if original_tokens == masked_tokens:
        return True

    original_base = strip_optional_business_tokens(original_tokens)
    masked_base = strip_optional_business_tokens(masked_tokens)
    return bool(original_base and masked_base and original_base == masked_base)


def canonical_entity_tokens(value: str) -> list[str]:
    normalized = value.upper()
    normalized = re.sub(r"\bM\s*/\s*S\b", " MS ", normalized)
    normalized = re.sub(r"\bM\s*\.\s*S\b", " MS ", normalized)
    normalized = normalized.replace("&", " AND ")
    normalized = re.sub(r"[^A-Z0-9 ]+", " ", normalized)
    raw_tokens = [token for token in re.sub(r"\s+", " ", normalized).strip().split() if token]

    tokens: list[str] = []
    index = 0
    while index < len(raw_tokens):
        token = raw_tokens[index]
        next_token = raw_tokens[index + 1] if index + 1 < len(raw_tokens) else ""
        third_token = raw_tokens[index + 2] if index + 2 < len(raw_tokens) else ""

        if index == 0 and token == "M" and next_token == "S":
            tokens.append("MS")
            index += 2
            continue
        if index == 0 and token == "MS":
            tokens.append("MS")
            index += 1
            continue
        if token == "R" and next_token == "M" and third_token == "C":
            tokens.append("RMC")
            index += 3
            continue
        if token in {"PRIVATE", "PVT", "PVTLTD"}:
            if token == "PVTLTD":
                tokens.append("PRIVATE_LIMITED")
                index += 1
                continue
            if next_token in {"LIMITED", "LTD", "LTD."}:
                tokens.append("PRIVATE_LIMITED")
                index += 2
                continue
        if token == "LIMIT" and next_token == "":
            tokens.append("PRIVATE_LIMITED")
            index += 1
            continue
        if token == "LTD":
            tokens.append("LIMITED")
            index += 1
            continue
        tokens.append(token)
        index += 1

    return tokens


def strip_optional_business_tokens(tokens: list[str]) -> list[str]:
    stripped = list(tokens)
    if stripped and stripped[0] == "MS":
        stripped = stripped[1:]
    while stripped and stripped[-1] in {"PRIVATE_LIMITED", "LIMITED"}:
        stripped = stripped[:-1]
    return [token for token in stripped if token != "AND"]


def format_literal_output(value: str) -> str:
    output = value.upper().strip()
    acronym_placeholders: dict[str, str] = {}

    def protect_acronym(match: re.Match[str]) -> str:
        key = f"__ACRONYM_{len(acronym_placeholders)}__"
        acronym_placeholders[key] = match.group(0)
        return key

    output = re.sub(r"\b(?:[A-Z]\.){2,}", protect_acronym, output)
    output = re.sub(r"\bM\s*/\s*S\b", "M/S", output)
    output = re.sub(r"\bM\s*\.\s*S\b", "M/S", output)
    output = output.replace("&", " & ")
    output = re.sub(r"\bPVT\s*\.\s*LTD\s*\.?\b", "PVT LTD", output)
    output = re.sub(r"\bPVT\s*\.\b", "PVT", output)
    output = re.sub(r"\bLTD\s*\.\b", "LTD", output)
    output = re.sub(r"[.]+", " ", output)
    for key, acronym in acronym_placeholders.items():
        output = output.replace(key, acronym)
    output = re.sub(r"\s+", " ", output).strip()
    return output


def build_literal_match_details(original: str, masked: str, rule: str) -> str:
    masked_tokens = tokenize_name(masked, keep_stars=True)
    stars = sum(token.count("*") for token in masked_tokens)
    visible = "".join(token.replace("*", "") for token in masked_tokens)
    first_visible = visible[0] if visible else ""
    last_visible = visible[-1] if visible else ""
    return (
        f"stars={stars}; first={first_visible or '-'}; last={last_visible or '-'}; "
        f"masked_words={len(masked_tokens)}; original_words={len(tokenize_name(original, keep_stars=False))}; "
        f"canonical_original={' '.join(canonical_entity_tokens(original))}; "
        f"canonical_masked={' '.join(canonical_entity_tokens(masked))}; rule={rule}"
    )


def greedy_span_matches(masked_tokens: list[str], original_tokens: list[str]) -> list[SpanMatch]:
    candidates: list[SpanMatch] = []
    max_span_size = min(3, len(original_tokens))

    for masked_index, masked_token in enumerate(masked_tokens):
        for original_start in range(len(original_tokens)):
            for original_end in range(original_start + 1, min(len(original_tokens), original_start + max_span_size) + 1):
                span_tokens = original_tokens[original_start:original_end]
                output_token = "".join(span_tokens) if len(span_tokens) > 1 else span_tokens[0]
                score = score_masked_token(masked_token, output_token)
                if len(span_tokens) > 1 and len(compact(masked_token, keep_stars=True)) != len(compact(output_token)):
                    score = min(score, 45.0)
                candidates.append(
                    SpanMatch(
                        masked_index=masked_index,
                        original_start=original_start,
                        original_end=original_end,
                        score=score,
                        output_token=output_token,
                    )
                )

    selected: list[SpanMatch] = []
    used_masked: set[int] = set()
    used_original: set[int] = set()
    for candidate in sorted(candidates, key=lambda item: item.score, reverse=True):
        original_indexes = set(range(candidate.original_start, candidate.original_end))
        if candidate.masked_index in used_masked or original_indexes & used_original:
            continue
        selected.append(candidate)
        used_masked.add(candidate.masked_index)
        used_original.update(original_indexes)
    return selected


def score_masked_token(masked_token: str, original_token: str) -> float:
    masked = compact(masked_token, keep_stars=True)
    original = compact(original_token)
    visible = masked.replace("*", "")

    if not masked or not original:
        return 0.0
    if masked == original:
        return 100.0
    if canonical_token_alias(masked) == canonical_token_alias(original):
        return 100.0
    if not visible:
        star_density = masked.count("*") / max(len(masked), 1)
        return 18.0 if star_density > 0.5 else 5.0

    order_ok = is_subsequence(visible, original)
    pattern_ok = masked_pattern_matches(masked, original)
    first_ok = visible[0] == original[0]
    last_ok = visible[-1] == original[-1]
    length_score = flexible_length_score(masked, original)
    visible_strength = min(12.0, (len(visible) / max(len(original), 1)) * 30.0)
    fuzzy_score = max(fuzzy_ratio(visible, original), fuzzy_partial_ratio(visible, original))

    score = 0.0
    score += 30.0 if pattern_ok else 0.0
    score += 18.0 if order_ok else 0.0
    score += 14.0 if first_ok else 0.0
    score += 14.0 if last_ok else 0.0
    score += length_score * 0.10
    score += visible_strength
    score += fuzzy_score * 0.12

    if len(visible) == 1 and not (first_ok or last_ok):
        score *= 0.55
    if not order_ok and not first_ok and not last_ok:
        score *= 0.45

    return min(100.0, score)


def canonical_token_alias(token: str) -> str:
    clean = compact(token)
    aliases = {
        "PVT": "PRIVATE",
        "PT": "PRIVATE",
        "PVTLTD": "PRIVATE_LIMITED",
        "PRIVATE": "PRIVATE",
        "LTD": "LIMITED",
        "LIMIT": "LIMITED",
        "LIMITED": "LIMITED",
        "MS": "MS",
    }
    return aliases.get(clean, clean)


def masked_pattern_matches(masked: str, original: str) -> bool:
    parts: list[str] = []
    for char in masked:
        if char == "*":
            if not parts or parts[-1] != "[A-Z0-9]*":
                parts.append("[A-Z0-9]*")
        else:
            parts.append(re.escape(char))
    pattern = "^" + "".join(parts) + "$"
    return re.match(pattern, original) is not None


def is_subsequence(needle: str, haystack: str) -> bool:
    cursor = 0
    for char in haystack:
        if cursor < len(needle) and needle[cursor] == char:
            cursor += 1
    return cursor == len(needle)


def flexible_length_score(masked: str, original: str) -> float:
    visible_count = len(masked.replace("*", ""))
    if visible_count > len(original):
        return 0.0
    if len(masked) == len(original):
        return 100.0
    distance = abs(len(masked) - len(original))
    return max(0.0, 100.0 - (distance / max(len(original), 1)) * 100.0)


def row_visible_similarity(original: str, masked: str) -> float:
    original_clean = compact(original)
    masked_visible = compact(masked, keep_stars=True).replace("*", "")
    if not original_clean or not masked_visible:
        return 0.0
    order_bonus = 20.0 if is_subsequence(masked_visible, original_clean) else 0.0
    fuzzy_score = max(fuzzy_ratio(masked_visible, original_clean), fuzzy_partial_ratio(masked_visible, original_clean))
    return min(100.0, fuzzy_score * 0.78 + order_bonus)


def word_count_similarity(original_count: int, masked_count: int) -> float:
    if original_count == 0 and masked_count == 0:
        return 100.0
    distance = abs(original_count - masked_count)
    return max(0.0, 100.0 - (distance / max(original_count, masked_count, 1)) * 100.0)


def initials_similarity(original_tokens: list[str], masked_tokens: list[str]) -> float:
    original_initials = [token[0] for token in original_tokens if token]
    masked_initials = [token.replace("*", "")[0] for token in masked_tokens if token.replace("*", "")]
    if not original_initials or not masked_initials:
        return 0.0
    hits = sum(1 for initial in masked_initials if initial in original_initials)
    return (hits / max(len(masked_initials), 1)) * 100.0


def count_unmatched_specific_terms(original_tokens: list[str], masked_tokens: list[str]) -> int:
    unmatched = 0
    for masked_token in masked_tokens:
        if masked_token == "&":
            continue
        masked = compact(masked_token, keep_stars=True)
        visible = masked.replace("*", "")
        if len(visible) < 2:
            continue
        _, score = best_masked_token_output(masked_token, original_tokens)
        if score < 68:
            unmatched += 1
    return unmatched


def count_matchable_tokens(tokens: list[str]) -> int:
    return sum(1 for token in tokens if token != "&")


def is_ms_prefix_token(token: str) -> bool:
    normalized = token.upper().replace("\\", "/")
    return normalized in {"M/S", "M/*", "M/"} or re.fullmatch(r"M\s*/\s*[*S]", normalized) is not None


def is_smt_prefix_token(token: str) -> bool:
    clean = compact(token, keep_stars=True)
    visible = clean.replace("*", "")
    return "*" in clean and len(clean) <= 3 and visible in {"S", "ST", "SMT"}


def normalize_prefix_output(masked_tokens: list[str], output: str) -> str:
    if not output:
        return output
    if masked_tokens and is_ms_prefix_token(masked_tokens[0]):
        output = re.sub(r"^M\s+S\b", "M/S", output)
        output = re.sub(r"^MS\b", "M/S", output)
    return output


def reconstruct_masked_output(original_tokens: list[str], masked_tokens: list[str]) -> Reconstruction:
    output_tokens: list[str] = []
    scores: list[float] = []
    matched_count = 0
    strong_seen = False
    index = 0

    while index < len(masked_tokens):
        masked_token = masked_tokens[index]
        if masked_token == "&":
            output_tokens.append("&")
            scores.append(100.0)
            index += 1
            continue

        masked_compact = compact(masked_token, keep_stars=True)
        visible = masked_compact.replace("*", "")

        if is_ms_prefix_token(masked_token):
            output_tokens.append("M/S")
            scores.append(96.0)
            matched_count += 1
            strong_seen = True
            index += 1
            continue

        if (
            index + 1 < len(masked_tokens)
            and compact(masked_tokens[index]) == "M"
            and compact(masked_tokens[index + 1]) == "S"
        ):
            output_tokens.append("M/S")
            scores.extend([92.0, 92.0])
            matched_count += 2
            strong_seen = True
            index += 2
            continue

        if (
            index == 0
            and masked_compact.startswith("M")
            and "*" in masked_compact
            and len(masked_compact) <= 2
            and len(masked_tokens) > 1
        ):
            output_tokens.append("Mr.")
            scores.append(88.0)
            matched_count += 1
            strong_seen = True
            index += 1
            continue

        if (
            index == 0
            and is_smt_prefix_token(masked_token)
            and len(masked_tokens) > 1
        ):
            output_tokens.append("SMT")
            scores.append(88.0)
            matched_count += 1
            strong_seen = True
            index += 1
            continue

        initial_run = collect_initial_run(masked_tokens, index)
        if len(initial_run) > 1 and "".join(initial_run) in {compact(token) for token in original_tokens}:
            output_tokens.extend(initial_run)
            scores.extend([96.0] * len(initial_run))
            matched_count += len(initial_run)
            strong_seen = True
            index += len(initial_run)
            continue

        candidate, score = best_masked_token_output(masked_token, original_tokens)
        if candidate and score >= 68:
            output_tokens.append(candidate)
            scores.append(score)
            matched_count += 1
            strong_seen = True
        elif visible and (len(visible) <= 2 or strong_seen):
            output_tokens.append(visible if "*" not in masked_compact else masked_token)
            scores.append(58.0 if strong_seen else 0.0)
        index += 1

    if not strong_seen:
        return Reconstruction("", 0, 0.0)

    return Reconstruction(" ".join(output_tokens), matched_count, average(scores))


def collect_initial_run(masked_tokens: list[str], start_index: int) -> list[str]:
    initials: list[str] = []
    for token in masked_tokens[start_index:]:
        clean = compact(token, keep_stars=True)
        if "*" in clean or len(clean) != 1:
            break
        initials.append(clean)
    return initials


def best_masked_token_output(masked_token: str, original_tokens: list[str]) -> tuple[str, float]:
    masked = compact(masked_token, keep_stars=True)
    if not masked:
        return "", 0.0

    candidates: list[tuple[str, float]] = []
    max_span_size = min(4, len(original_tokens))
    for original_start in range(len(original_tokens)):
        for original_end in range(original_start + 1, min(len(original_tokens), original_start + max_span_size) + 1):
            span_tokens = original_tokens[original_start:original_end]
            span_compact = "".join(compact(token) for token in span_tokens)
            display = format_span_output(span_tokens)
            span_score = score_masked_token(masked, span_compact)
            if masked.endswith("S") and not span_compact.endswith("S"):
                span_score = max(span_score, score_masked_token(masked, f"{span_compact}S") - 2)
            if len(span_tokens) > 1 and len(masked) != len(span_compact):
                span_score = min(span_score, 45.0)
            candidates.append((display, span_score))

            masked_length = len(masked)
            if "*" in masked and len(span_tokens) == 1 and 0 < masked_length <= len(span_compact):
                for start in range(0, len(span_compact) - masked_length + 1):
                    segment = span_compact[start : start + masked_length]
                    segment_score = score_masked_token(masked, segment)
                    if segment_score >= 68:
                        candidates.append((display_segment(masked, segment), segment_score))

    if not candidates:
        return "", 0.0

    output, score = max(candidates, key=lambda item: item[1])
    return output, score


def format_span_output(tokens: list[str]) -> str:
    if len(tokens) <= 1:
        return display_business_alias(tokens[0]) if tokens else ""
    if len(tokens) == 2 and all(len(token) <= 5 for token in tokens):
        return "".join(tokens)
    return " ".join(display_business_alias(token) for token in tokens)


def display_segment(masked: str, segment: str) -> str:
    clean_segment = compact(segment)
    visible = masked.replace("*", "")
    if clean_segment == "PRIVATE" and visible in {"P", "PT", "PVT"}:
        return "PRIVATE"
    if clean_segment == "LIMITED" and len(visible) <= 1:
        return "LTD"
    return clean_segment


def display_business_alias(token: str) -> str:
    clean = compact(token)
    if clean == "PRIVATE":
        return "PRIVATE"
    if clean == "LIMITED":
        return "LIMITED"
    return token


def reconstruct_partial_output(
    original_tokens: list[str],
    masked_tokens: list[str],
    strong_matches: list[SpanMatch],
    all_matches: list[SpanMatch],
) -> str:
    if strong_matches:
        matched_masked_count = len({match.masked_index for match in strong_matches})
        ordered_matches = sorted(strong_matches, key=lambda match: match.masked_index)
        first_match = min(match.original_start for match in strong_matches)

        if matched_masked_count == len(masked_tokens):
            if first_match > 0:
                return " ".join(original_tokens[:first_match])
            return " ".join(match.output_token for match in ordered_matches)

        first_ordered = ordered_matches[0]
        if first_ordered.original_start == 0:
            return original_tokens[0]
        return first_ordered.output_token

    if all_matches:
        best_match = max(all_matches, key=lambda match: match.score)
        return original_tokens[best_match.original_start]

    return " ".join(original_tokens[:1])


def reconstruct_masked_order_output(
    masked_tokens: list[str],
    strong_matches: list[SpanMatch],
) -> str:
    match_by_masked_index = {match.masked_index: match for match in strong_matches}
    ordered_tokens: list[str] = []
    used_original_indexes: set[int] = set()

    for masked_index in range(len(masked_tokens)):
        match = match_by_masked_index.get(masked_index)
        if match is None:
            continue
        original_indexes = set(range(match.original_start, match.original_end))
        if original_indexes & used_original_indexes:
            continue
        ordered_tokens.append(match.output_token)
        used_original_indexes.update(original_indexes)

    return " ".join(ordered_tokens)


def build_match_details(
    original_tokens: list[str],
    masked_tokens: list[str],
    matches: list[SpanMatch],
    rule: str,
) -> str:
    stars = sum(token.count("*") for token in masked_tokens)
    visible = "".join(token.replace("*", "") for token in masked_tokens)
    first_visible = visible[0] if visible else ""
    last_visible = visible[-1] if visible else ""

    pair_parts: list[str] = []
    for match in sorted(matches, key=lambda item: item.masked_index):
        if match.masked_index >= len(masked_tokens) or match.original_start >= len(original_tokens):
            continue
        pair_parts.append(
            f"{masked_tokens[match.masked_index]}->{match.output_token}({clamp_int(match.score)}%)"
        )

    matched = ", ".join(pair_parts) if pair_parts else "none"
    return (
        f"stars={stars}; first={first_visible or '-'}; last={last_visible or '-'}; "
        f"masked_words={len(masked_tokens)}; original_words={len(original_tokens)}; "
        f"matches={matched}; rule={rule}"
    )


def average(values: Iterable[float]) -> float:
    values = list(values)
    return sum(values) / len(values) if values else 0.0


def clamp_int(value: float) -> int:
    return int(round(max(0.0, min(100.0, value))))


def autosize_columns(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    worksheet = writer.sheets[sheet_name]
    for index, column in enumerate(df.columns, start=1):
        values = df[column].head(200).tolist()
        max_length = max([len(str(column)), *(len(str(value)) for value in values)])
        worksheet.column_dimensions[worksheet.cell(row=1, column=index).column_letter].width = min(
            max(max_length + 2, 14), 45
        )
