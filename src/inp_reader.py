"""Utilities for reading and summarizing SACS text input files."""

from __future__ import annotations

import json
import os
import re
from collections import Counter
from dataclasses import dataclass
from typing import Any


PREFERRED_ENCODINGS = ("utf-8", "utf-8-sig", "gb18030", "gbk", "latin-1")
BASIC_LOAD_CASE_TITLE = "** SEASTATE BASIC LOAD CASE DESCRIPTIONS **"
BASIC_LOAD_CASE_MARKER = "LOAD  LOAD     ********** DESCRIPTION ***********"
BASIC_LOAD_CASE_SUBHEADER = "CASE  LABEL"
BASIC_LOAD_CASE_STOP_MARKER = "SEASTATE BASIC LOAD CASE SUMMARY"


@dataclass
class SACSInpDocument:
    path: str
    encoding: str
    text: str

    @property
    def lines(self) -> list[str]:
        return self.text.splitlines()

    @property
    def non_empty_lines(self) -> list[str]:
        return [line for line in self.lines if line.strip()]


class SACSInpReader:
    def __init__(self, path: str):
        self.path = path

    def read(self) -> SACSInpDocument:
        if not os.path.exists(self.path):
            raise FileNotFoundError(f"未找到 SACS 输入文件: {self.path}")

        with open(self.path, "rb") as file:
            raw_bytes = file.read()

        for encoding in PREFERRED_ENCODINGS:
            try:
                text = raw_bytes.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            text = raw_bytes.decode("latin-1", errors="replace")
            encoding = "latin-1"

        normalized_text = text.replace("\r\n", "\n").replace("\r", "\n")
        return SACSInpDocument(path=self.path, encoding=encoding, text=normalized_text)

    def summarize(self, preview_limit: int = 12, top_card_limit: int = 12) -> dict[str, Any]:
        document = self.read()
        card_counter = Counter()
        for line in document.non_empty_lines:
            card_counter[self._card_name(line)] += 1

        basic_load_case_lines = self.extract_basic_load_case_descriptions(document)

        return {
            "path": document.path,
            "file_name": os.path.basename(document.path),
            "encoding": document.encoding,
            "total_lines": len(document.lines),
            "non_empty_lines": len(document.non_empty_lines),
            "preview_lines": document.non_empty_lines[:preview_limit],
            "basic_load_case_count": max(len(basic_load_case_lines) - 3, 0),
            "basic_load_case_lines": basic_load_case_lines,
            "top_cards": [
                {"card": card, "count": count}
                for card, count in card_counter.most_common(top_card_limit)
            ],
        }

    def extract_basic_load_case_descriptions(
        self,
        document: SACSInpDocument | None = None,
    ) -> list[str]:
        document = document or self.read()
        extracted_lines: list[str] = []
        has_started = False

        for raw_line in document.lines:
            normalized_line = raw_line.replace("\f", "").rstrip()
            stripped_line = normalized_line.strip()

            if BASIC_LOAD_CASE_STOP_MARKER in stripped_line and has_started:
                break

            if BASIC_LOAD_CASE_MARKER in stripped_line:
                if not extracted_lines:
                    extracted_lines.extend(
                        [
                            BASIC_LOAD_CASE_TITLE,
                            BASIC_LOAD_CASE_MARKER,
                            BASIC_LOAD_CASE_SUBHEADER,
                        ]
                    )
                has_started = True
                continue

            if not has_started or not stripped_line:
                continue

            if BASIC_LOAD_CASE_TITLE in stripped_line:
                continue
            if stripped_line == BASIC_LOAD_CASE_SUBHEADER:
                continue
            if stripped_line.startswith("SACS CONNECT Edition"):
                continue
            if "SACS IV SEASTATE PROGRAM" in stripped_line:
                continue

            if self._is_basic_load_case_data_line(stripped_line):
                extracted_lines.append(normalized_line)

        return extracted_lines

    @staticmethod
    def _card_name(line: str) -> str:
        stripped = line.strip()
        if not stripped:
            return "BLANK"
        if stripped.startswith(("*", "!", "$")):
            return "COMMENT"

        fixed_width_card = line[:4].strip().upper()
        if fixed_width_card.isalpha():
            return fixed_width_card

        first_token = stripped.split()[0].upper()
        return first_token[:12]

    @staticmethod
    def _is_basic_load_case_data_line(line: str) -> bool:
        return re.match(r"^\d+\s+\S+\s+.+$", line) is not None


def _main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Read and summarize a SACS input file")
    parser.add_argument("path", help="Path to the SACS input file")
    parser.add_argument("--preview", type=int, default=12, help="Number of preview lines")
    parser.add_argument("--cards", type=int, default=12, help="Number of top cards")
    args = parser.parse_args()

    summary = SACSInpReader(args.path).summarize(
        preview_limit=args.preview,
        top_card_limit=args.cards,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
