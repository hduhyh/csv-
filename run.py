"""Command-line entry point for CSV/Excel merging."""

from __future__ import annotations

import argparse

from merge_core import merge_csv_files


def main() -> None:
    parser = argparse.ArgumentParser(description="按列名合并文件夹中的 CSV 和 Excel 文件")
    parser.add_argument("input_folder", nargs="?", default="./input", help="输入文件夹")
    parser.add_argument("-o", "--output-folder", help="输出文件夹，默认与输入文件夹相同")
    parser.add_argument("-t", "--template", help="列名模板（CSV/XLSX/XLSM）")
    parser.add_argument("-n", "--name", default="merged_output.csv", help="输出文件名")
    parser.add_argument("--no-source", action="store_true", help="不添加 Source_File 来源列")
    args = parser.parse_args()

    result = merge_csv_files(
        input_folder=args.input_folder,
        output_folder=args.output_folder,
        template_path=args.template,
        output_filename=args.name,
        add_source_column=not args.no_source,
        progress_callback=print,
    )
    print(f"输出文件：{result.output_path}")
    if result.warning_path:
        print(f"检查报告：{result.warning_path}")


if __name__ == "__main__":
    main()
