"""KOFIA 자산운용사 종합 프로파일 CLI."""

import argparse
import os
import sys
import time

from .collector import (
    FreesisClient,
    bootstrap_session,
    build_column_map,
    collect_freesis_all_modes,
    fetch_metadata,
    save_freesis_csv,
)
from .joiner import build_joined_excel, build_joined_excel_from_csvs


REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_RAW = os.path.join(REPO_ROOT, "data", "raw")
DATA_PROCESSED = os.path.join(REPO_ROOT, "data", "processed")


def cmd_run(args):
    print("=" * 60)
    print("KOFIA 자산운용사 종합 프로파일 V4")
    print("=" * 60)

    date_override = getattr(args, "date", None)

    print("\n[1/4] FreeSIS 데이터 수집...")
    client = FreesisClient()
    bootstrap_session(client)
    time.sleep(0.5)

    metadata = fetch_metadata(client)
    if not metadata:
        print("FATAL: FreeSIS 메타데이터 수집 실패")
        sys.exit(1)

    ordered_keys, key_to_label = build_column_map(metadata)
    print(f"  컬럼 매핑: {len(ordered_keys)}개 (ORDER_SEQ 순서)")

    freesis_results = collect_freesis_all_modes(
        client, metadata, ordered_keys, key_to_label, date_override=date_override
    )

    for label, result in freesis_results.items():
        print(f"  {label}: {len(result['rows'])} rows")

    print("\n[2/4] RAW 데이터 저장...")
    os.makedirs(DATA_RAW, exist_ok=True)
    save_freesis_csv(freesis_results, DATA_RAW, ordered_keys)

    print("\n[3/4] 조인 + 엑셀 생성...")
    os.makedirs(DATA_PROCESSED, exist_ok=True)
    output_path = os.path.join(DATA_PROCESSED, "KOFIA_자산운용사_종합프로파일.xlsx")
    build_joined_excel(freesis_results, ordered_keys, key_to_label, metadata, output_path)

    print("\n[4/4] 완료!")
    print(f"  출력: {output_path}")


def cmd_join(args):
    print("join 명령은 V4에서 지원하지 않습니다. 'run' 또는 'rebuild'을 사용하세요.")
    sys.exit(1)


def cmd_rebuild(args):
    print("=" * 60)
    print("KOFIA 자산운용사 종합 프로파일 V4 - REBUILD")
    print("=" * 60)

    print("\n[1/2] CSV에서 엑셀 재생성...")
    os.makedirs(DATA_PROCESSED, exist_ok=True)
    output_path = os.path.join(DATA_PROCESSED, "KOFIA_자산운용사_종합프로파일.xlsx")
    build_joined_excel_from_csvs(DATA_RAW, output_path)

    print("\n[2/2] 완료!")
    print(f"  출력: {output_path}")
    print("  (data/raw/ CSV 원본은 변경 없음)")


def main():
    parser = argparse.ArgumentParser(description="KOFIA 자산운용사 종합 프로파일")
    sub = parser.add_subparsers(dest="command")

    run_parser = sub.add_parser("run", help="전체 파이프라인: 수집 + 조인 + 엑셀 생성")
    run_parser.add_argument("--date", type=str, default=None, help="기준일 오버라이드 (예: 20260916)")

    sub.add_parser("join", help="더 이상 지원하지 않음")

    rebuild_parser = sub.add_parser("rebuild", help="CSV에서 엑셀 재생성 (API 불필요)")

    args = parser.parse_args()
    if args.command == "run":
        cmd_run(args)
    elif args.command == "join":
        cmd_join(args)
    elif args.command == "rebuild":
        cmd_rebuild(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
