import argparse
from scripts.parse_all import parse_all_documents


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Document Parser CLI"
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    parse_cmd = subparsers.add_parser("parse", help="Parse documents")
    parse_cmd.add_argument("--input", required=True, help="Input directory")
    parse_cmd.add_argument("--output", required=True, help="Output directory")
    parse_cmd.add_argument(
        "--recursive",
        action="store_true",
        help="Recursively scan subfolders",
    )
    parse_cmd.add_argument(
        "--ext",
        default="pdf,docx",
        help="Comma-separated file extensions",
    )
    parse_cmd.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )

    args = parser.parse_args()

    if args.command == "parse":

        parse_all_documents(
            input_directory=args.input,
            output_directory=args.output,
        )


if __name__ == "__main__":
    main()
