import argparse
from src import config


def setup_args():
    parser = argparse.ArgumentParser(
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    parser.add_argument(
        "--file",
        dest="template_path",
        default=config.TEMPLATE_PATH,
        help="Custom template file path"
    )
    return parser.parse_args()


def apply_args(args):
    """Apply CLI args to runtime config to single source of truth"""
    if args.template_path:
        config.TEMPLATE_PATH = args.template_path