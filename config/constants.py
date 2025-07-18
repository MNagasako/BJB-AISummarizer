import os

BASE_DIR = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
INPUT_DIR = os.path.join(BASE_DIR, "input")
PRIVATE_INPUT_DIR = os.path.join(INPUT_DIR, ".private")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
CONFIG_DIR = os.path.join(BASE_DIR, "config")
PROMPT_DIR = os.path.join(BASE_DIR, "prompts")

DEFAULT_OUTPUT_MAP = os.path.join(CONFIG_DIR, "output_map.json")
DEFAULT_AI_CONFIG = os.path.join(CONFIG_DIR, "ai_config.json")

PROMPT_PLACEHOLDER_PATTERN = r"%([^ -%]+)%"
