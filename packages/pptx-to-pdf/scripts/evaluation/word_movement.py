#!/usr/bin/env python3
# Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
# Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
# Production use requires a commercial agreement: licensing@eigenpal.com
#
# Adapted from DOCX Editor `packages/docx-to-pdf/scripts/pdf-visual-diff.py`, commit
# 398963c7f67207cebf3c38dcc41d5f0bd1339b5f: `normalized_word`, `word_matches`, and
# `compare_word_movement`, unchanged. The rest of that file (its rendering command line) is not
# copied; `worker.py` needs only the word movement.
"""Word alignment and movement between a reference and a candidate PDF."""

from __future__ import annotations

import difflib
import math
import statistics
import unicodedata
from typing import Any


def normalized_word(text: str) -> str:
    return " ".join(unicodedata.normalize("NFC", text).split())


def word_matches(
    reference: list[dict[str, Any]],
    candidate: list[dict[str, Any]],
) -> list[tuple[int, int]]:
    def monotonic_geometry_pairs(
        left_words: list[dict[str, Any]],
        left_indices: list[int],
        right_words: list[dict[str, Any]],
        right_indices: list[int],
    ) -> list[tuple[int, int]]:
        if len(left_indices) > len(right_indices):
            return [
                (right_index, left_index)
                for left_index, right_index in monotonic_geometry_pairs(
                    right_words,
                    right_indices,
                    left_words,
                    left_indices,
                )
            ]
        if len(left_indices) * len(right_indices) > 100_000:
            return list(zip(left_indices, right_indices))
        rows = len(left_indices)
        columns = len(right_indices)
        costs = [[math.inf] * (columns + 1) for _ in range(rows + 1)]
        matched = [[False] * (columns + 1) for _ in range(rows + 1)]
        for column in range(columns + 1):
            costs[0][column] = 0
        for row in range(1, rows + 1):
            left = left_words[left_indices[row - 1]]
            left_x = (left["x0"] + left["x1"]) / 2
            left_y = (left["y0"] + left["y1"]) / 2
            for column in range(1, columns + 1):
                skip_cost = costs[row][column - 1]
                right = right_words[right_indices[column - 1]]
                right_x = (right["x0"] + right["x1"]) / 2
                right_y = (right["y0"] + right["y1"]) / 2
                pair_cost = costs[row - 1][column - 1] + math.hypot(
                    right_x - left_x,
                    right_y - left_y,
                )
                if pair_cost < skip_cost:
                    costs[row][column] = pair_cost
                    matched[row][column] = True
                else:
                    costs[row][column] = skip_cost
        pairs: list[tuple[int, int]] = []
        row = rows
        column = columns
        while row > 0 and column > 0:
            if matched[row][column]:
                pairs.append((left_indices[row - 1], right_indices[column - 1]))
                row -= 1
                column -= 1
            else:
                column -= 1
        pairs.reverse()
        return pairs

    matches: list[tuple[int, int]] = []
    matched_reference: set[int] = set()
    matched_candidate: set[int] = set()
    pages = sorted({word["page"] for word in reference} | {word["page"] for word in candidate})

    for page in pages:
        reference_indices = [
            index for index, word in enumerate(reference) if word["page"] == page
        ]
        candidate_indices = [
            index for index, word in enumerate(candidate) if word["page"] == page
        ]
        matcher = difflib.SequenceMatcher(
            None,
            [reference[index]["text"] for index in reference_indices],
            [candidate[index]["text"] for index in candidate_indices],
            autojunk=False,
        )
        for block in matcher.get_matching_blocks():
            for offset in range(block.size):
                reference_index = reference_indices[block.a + offset]
                candidate_index = candidate_indices[block.b + offset]
                matches.append((reference_index, candidate_index))
                matched_reference.add(reference_index)
                matched_candidate.add(candidate_index)

    def pair_residuals(same_page: bool) -> None:
        reference_groups: dict[tuple[Any, ...], list[int]] = {}
        candidate_groups: dict[tuple[Any, ...], list[int]] = {}
        for index, word in enumerate(reference):
            if index in matched_reference:
                continue
            key = (word["page"], word["text"]) if same_page else (word["text"],)
            reference_groups.setdefault(key, []).append(index)
        for index, word in enumerate(candidate):
            if index in matched_candidate:
                continue
            key = (word["page"], word["text"]) if same_page else (word["text"],)
            candidate_groups.setdefault(key, []).append(index)
        for key in sorted(reference_groups.keys() & candidate_groups.keys()):
            reference_indices = sorted(
                reference_groups[key],
                key=lambda index: (
                    reference[index]["page"],
                    reference[index]["y0"],
                    reference[index]["x0"],
                    index,
                ),
            )
            candidate_indices = sorted(
                candidate_groups[key],
                key=lambda index: (
                    candidate[index]["page"],
                    candidate[index]["y0"],
                    candidate[index]["x0"],
                    index,
                ),
            )
            for reference_index, candidate_index in zip(reference_indices, candidate_indices):
                matches.append((reference_index, candidate_index))
                matched_reference.add(reference_index)
                matched_candidate.add(candidate_index)

    pair_residuals(True)
    pair_residuals(False)

    # Sequence alignment can pair repeated labels across distant table rows when nearby
    # text wraps differently. Equal same-page inventories retain reading-order identity.
    repeated_reference: dict[tuple[int, str], list[int]] = {}
    repeated_candidate: dict[tuple[int, str], list[int]] = {}
    for index, word in enumerate(reference):
        repeated_reference.setdefault((word["page"], word["text"]), []).append(index)
    for index, word in enumerate(candidate):
        repeated_candidate.setdefault((word["page"], word["text"]), []).append(index)
    for key in repeated_reference.keys() & repeated_candidate.keys():
        reference_indices = repeated_reference[key]
        candidate_indices = repeated_candidate[key]
        if len(reference_indices) < 2 or len(candidate_indices) < 2:
            continue
        reference_set = set(reference_indices)
        candidate_set = set(candidate_indices)
        matches = [
            pair
            for pair in matches
            if pair[0] not in reference_set and pair[1] not in candidate_set
        ]
        reading_order = lambda words, index: (
            words[index]["y0"],
            words[index]["x0"],
            index,
        )
        ordered_reference = sorted(
            reference_indices,
            key=lambda index: reading_order(reference, index),
        )
        ordered_candidate = sorted(
            candidate_indices,
            key=lambda index: reading_order(candidate, index),
        )
        for reference_index, candidate_index in monotonic_geometry_pairs(
            reference,
            ordered_reference,
            candidate,
            ordered_candidate,
        ):
            matches.append((reference_index, candidate_index))

    matches.sort()
    return matches


def compare_word_movement(
    reference: list[dict[str, Any]],
    candidate: list[dict[str, Any]],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    matched_indices = word_matches(reference, candidate)
    matched_reference = {pair[0] for pair in matched_indices}
    matched_candidate = {pair[1] for pair in matched_indices}
    buckets = {
        "within0_5Pt": 0,
        "between0_5And2Pt": 0,
        "between2And8Pt": 0,
        "beyond8Pt": 0,
        "crossPage": 0,
    }
    distances: list[float] = []
    pairs: list[dict[str, Any]] = []
    weighted_penalty = 0.0
    for reference_index, candidate_index in matched_indices:
        reference_word = reference[reference_index]
        candidate_word = candidate[candidate_index]
        if reference_word["page"] != candidate_word["page"]:
            buckets["crossPage"] += 1
            weighted_penalty += 100
            distance = None
        else:
            distance = math.hypot(
                (candidate_word["x0"] + candidate_word["x1"]) / 2
                - (reference_word["x0"] + reference_word["x1"]) / 2,
                (candidate_word["y0"] + candidate_word["y1"]) / 2
                - (reference_word["y0"] + reference_word["y1"]) / 2,
            )
            distances.append(distance)
            if distance <= 0.5:
                buckets["within0_5Pt"] += 1
            elif distance <= 2:
                buckets["between0_5And2Pt"] += 1
            elif distance <= 8:
                buckets["between2And8Pt"] += 1
            else:
                buckets["beyond8Pt"] += 1
            weighted_penalty += min(100, distance * 5)
        pairs.append(
            {
                "text": reference_word["text"],
                "reference": reference_word,
                "candidate": candidate_word,
                "distancePt": round(distance, 4) if distance is not None else None,
            }
        )

    unmatched_reference = len(reference) - len(matched_reference)
    unmatched_candidate = len(candidate) - len(matched_candidate)
    if buckets["crossPage"]:
        severity = "critical"
    elif buckets["beyond8Pt"]:
        severity = "major"
    elif buckets["between2And8Pt"]:
        severity = "moderate"
    elif buckets["between0_5And2Pt"]:
        severity = "minor"
    else:
        severity = "equal"
    sorted_distances = sorted(distances)
    denominator = max(1, len(matched_indices) * 2)
    largest_movements = sorted(
        (
            {
                "text": pair["text"],
                "referencePage": pair["reference"]["page"],
                "candidatePage": pair["candidate"]["page"],
                "referenceXPt": round(pair["reference"]["x0"], 3),
                "referenceYPt": round(pair["reference"]["y0"], 3),
                "candidateXPt": round(pair["candidate"]["x0"], 3),
                "candidateYPt": round(pair["candidate"]["y0"], 3),
                "distancePt": pair["distancePt"],
            }
            for pair in pairs
            if pair["distancePt"] is None or pair["distancePt"] > 0.5
        ),
        key=lambda movement: (
            movement["distancePt"] is None,
            movement["distancePt"] or 0,
        ),
        reverse=True,
    )[:100]
    reading_order = sorted(pairs, key=lambda pair: (
        pair['reference']['page'], pair['reference']['y0'], pair['reference']['x0']))
    earliest_movements = [
        dict(text=pair['text'], referencePage=pair['reference']['page'],
             candidatePage=pair['candidate']['page'], referenceYPt=round(pair['reference']['y0'], 3),
             candidateYPt=round(pair['candidate']['y0'], 3), distancePt=pair['distancePt'])
        for pair in reading_order if pair['distancePt'] is None or pair['distancePt'] > 0.5
    ][:100]
    page_drift = []
    for page in sorted({pair['reference']['page'] for pair in reading_order}):
        same_page = [pair for pair in reading_order
                     if pair['reference']['page'] == page == pair['candidate']['page']]
        if not same_page:
            continue
        deltas = [((pair['candidate']['y0'] + pair['candidate']['y1'])
                   - (pair['reference']['y0'] + pair['reference']['y1'])) / 2 for pair in same_page]
        band = max(1, len(deltas) // 3)
        page_drift.append(dict(page=page, matchedTokens=len(deltas),
                               medianDeltaYPt=round(statistics.median(deltas), 4),
                               topThirdDeltaYPt=round(statistics.median(deltas[:band]), 4),
                               bottomThirdDeltaYPt=round(statistics.median(deltas[-band:]), 4)))
    summary = {
        "severity": severity,
        "referenceWords": len(reference),
        "candidateWords": len(candidate),
        "matchedWords": len(matched_indices),
        "unmatchedReferenceWords": unmatched_reference,
        "unmatchedCandidateWords": unmatched_candidate,
        "missingWordCount": unmatched_reference,
        "extraWordCount": unmatched_candidate,
        "distanceBuckets": buckets,
        "maxDistancePt": round(max(distances), 4) if distances else None,
        "medianDistancePt": (
            round(sorted_distances[len(sorted_distances) // 2], 4)
            if sorted_distances
            else None
        ),
        "movementScore": round(min(100, weighted_penalty / denominator), 4),
        "largestMovements": largest_movements,
        "earliestMovements": earliest_movements,
        "pageDrift": page_drift,
        "unmatchedReferenceSample": [
            reference[index]["text"]
            for index in range(len(reference))
            if index not in matched_reference
        ][:50],
        "unmatchedCandidateSample": [
            candidate[index]["text"]
            for index in range(len(candidate))
            if index not in matched_candidate
        ][:50],
    }
    return summary, pairs
