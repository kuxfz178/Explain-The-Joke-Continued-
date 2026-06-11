"""Generates run_mcq_openrouter.ipynb. Every model is called through OpenRouter."""
import nbformat as nbf

nb = nbf.v4.new_notebook()
cells = []

def md(text):
    cells.append(nbf.v4.new_markdown_cell(text))

def code(src):
    cells.append(nbf.v4.new_code_cell(src))


# ===== 1. Title =====
md("""# Picture MCQ Benchmark, OpenRouter version

This notebook runs every model through **OpenRouter**, which fronts the OpenAI, Anthropic, Google, Gemma, Qwen, and other providers behind one API.

One API key, one code path, no GPU needed. 

**Workflow:** set the model in the CONFIG cell, run every cell below it, then change the model and repeat. The run loop is resume-safe.
""")


# ===== 2. install =====
md("## 1. Install dependencies\n\nRun once per Colab session.")

code("""!pip install -q openai""")


# ===== 3. mount drive =====
md("""## 2. Mount Google Drive

Skip this if you are running locally on your Mac. When the prompt appears, authorise Colab to read your Drive.""")

code("""from google.colab import drive
drive.mount('/content/drive')""")


# ===== 4. paths =====
md("## 3. File paths\n\nEdit `PROJECT_DIR` so it points at your project folder on Drive.")

code("""import os

# CHANGE THIS to where your project files live on drive
PROJECT_DIR = "/content/drive/MyDrive/ExplainTheJoke"

INPUT_CSV    = f"{PROJECT_DIR}/model_input.csv"
METADATA_CSV = f"{PROJECT_DIR}/metadata.csv"
IMAGES_DIR   = f"{PROJECT_DIR}/images"
RUNS_DIR     = f"{PROJECT_DIR}/runs"

os.makedirs(RUNS_DIR, exist_ok=True)

# sanity check
for p, label in [(INPUT_CSV, "MCQ input"), (METADATA_CSV, "metadata"), (IMAGES_DIR, "images dir")]:
    print(f"{'OK ' if os.path.exists(p) else 'MISSING'}  {label}: {p}")""")


# ===== 5. API key =====
md("""## 4. OpenRouter API key

**Never paste the key directly into the notebook.** """)

code("""from google.colab import userdata
import os

os.environ["OPENROUTER_API_KEY"] = userdata.get("OPENROUTER_API_KEY")
print("loaded OPENROUTER_API_KEY")""")


# ===== 6. CONFIG =====
md("""## 5. CONFIG. Pick which model to run

**The only cell that changes between runs.** Set `MODEL`, run every cell below it, wait for it to finish, then change `MODEL` and run again.

Model names use OpenRouter's `provider/model` format. The full catalogue is at https://openrouter.ai/models.""")

code('''# Pick ONE of these.

# Closed models.
MODEL = "openai/gpt-5"
# MODEL = "anthropic/claude-opus-4-6"
# MODEL = "google/gemini-3.5-flash"

# Open-weight models. Same code path, OpenRouter just hosts them.
# MODEL = "google/gemma-3-4b-it"
# MODEL = "google/gemma-4-26b-a4b-it"
# MODEL = "qwen/qwen3-vl-30b-a3b-thinking"
# MODEL = "qwen/qwen3.5-9b"

# Output file is named after the model. Slashes and colons are replaced so the
# filename is safe on all platforms.
OUTPUT_CSV = f"{RUNS_DIR}/{MODEL.replace('/', '_').replace(':', '_')}_run.csv"

TEMPERATURE = 0.0
MAX_TOKENS  = 2000   # big enough to cover hidden reasoning tokens on the thinking models

print(f"will run: {MODEL}")
print(f"output:   {OUTPUT_CSV}")''')


# ===== 7. helpers =====
md("## 6. Shared helpers")

code('''import base64, csv, re, time
from pathlib import Path

LETTERS = ("A", "B", "C", "D")

# The prompt for every model. We deliberately keep it short and identical
# across providers, with one user turn (image plus four options).
PROMPT_TEMPLATE = """Look at the image and choose the option that best explains why it is funny.

Reply with exactly one letter: A, B, C, or D. No other text.

A) {a}
B) {b}
C) {c}
D) {d}"""


def encode_image_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def guess_media_type(path):
    return {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".webp": "image/webp", ".gif": "image/gif",
    }.get(Path(path).suffix.lower(), "image/jpeg")


def parse_letter(raw):
    """Pulls A, B, C, or D out of whatever string the model returned."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    if s.upper() in LETTERS:
        return s.upper()
    m = re.match(r"^[\\s*(\\[]*([ABCD])[\\s*).:\\]]", s.upper() + " ")
    if m:
        return m.group(1)
    hits = re.findall(r"\\b([ABCD])\\b", s.upper())
    if len(set(hits)) == 1:
        return hits[0]
    return None


def load_image_paths():
    paths = {}
    with open(METADATA_CSV, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            fid = (row.get("ID") or "").strip()
            rel = (row.get("img") or "").strip()
            if fid and rel:
                paths[fid] = os.path.join(IMAGES_DIR, os.path.basename(rel))
    return paths


def load_already_answered():
    done = set()
    if not os.path.exists(OUTPUT_CSV):
        return done
    with open(OUTPUT_CSV, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            fid = (row.get("File") or "").strip()
            if fid:
                done.add(fid)
    return done


id_to_image = load_image_paths()
print(f"loaded {len(id_to_image)} image paths")''')


# ===== 8. openrouter call =====
md("""## 7. The one model-call function

OpenRouter uses OpenAI's API shape, so we point the OpenAI SDK at OpenRouter's URL. The same function below handles every model.""")

code('''from openai import OpenAI

# Point the OpenAI SDK at OpenRouter instead of OpenAI.
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=os.environ["OPENROUTER_API_KEY"],
)


def call_openrouter(model, prompt, image_path):
    b64 = encode_image_base64(image_path)
    mime = guess_media_type(image_path)
    resp = client.chat.completions.create(
        model=model,
        temperature=TEMPERATURE,
        max_tokens=MAX_TOKENS,
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url",
                 "image_url": {"url": f"data:{mime};base64,{b64}"}},
            ],
        }],
        # Optional but recommended. OpenRouter uses these for analytics and to
        # fall back to a different host if the primary one is down.
        extra_headers={
            "HTTP-Referer": "https://github.com/",
            "X-Title": "Picture-MCQ benchmark",
        },
    )
    return resp.choices[0].message.content''')


# ===== 9. run loop =====
md("""## 8. Run the benchmark

Loops over `model_input.csv` and writes the answers to the output CSV. It is resume-safe, so if it crashes or Colab disconnects you can just rerun this cell and it picks up where it left off.

Set `LIMIT = 5` for a quick smoke test before committing to all 501 items.""")

code('''LIMIT = 5   # set to None for the full run. 5 is a quick smoke test (~30 seconds).

already_done = load_already_answered()
print(f"resuming. {len(already_done)} items already in {OUTPUT_CSV}")

is_new = not os.path.exists(OUTPUT_CSV) or os.path.getsize(OUTPUT_CSV) == 0
out_f = open(OUTPUT_CSV, "a", newline="", encoding="utf-8")
writer = csv.DictWriter(out_f, fieldnames=["File", "answer", "model"])
if is_new:
    writer.writeheader()

processed = invalid = errors = 0
SLEEP_SECONDS = 0.3

with open(INPUT_CSV, newline="", encoding="utf-8") as f:
    for row in csv.DictReader(f):
        fid = row["File"].strip()

        if fid in already_done:
            continue
        if fid not in id_to_image:
            print(f"  SKIP {fid}: not in metadata.csv")
            continue
        image_path = id_to_image[fid]
        if not os.path.exists(image_path):
            print(f"  SKIP {fid}: image missing at {image_path}")
            continue

        prompt = PROMPT_TEMPLATE.format(
            a=row["option_a"], b=row["option_b"],
            c=row["option_c"], d=row["option_d"],
        )

        try:
            raw = call_openrouter(MODEL, prompt, image_path)
            letter = parse_letter(raw)
            answer = letter or "INVALID"
            if letter is None:
                invalid += 1
        except Exception as e:
            print(f"  ERROR {fid}: {type(e).__name__}: {e}")
            answer = "ERROR"
            errors += 1

        writer.writerow({"File": fid, "answer": answer, "model": MODEL})
        out_f.flush()
        processed += 1

        if processed % 20 == 0:
            print(f"  processed {processed}  (invalid={invalid}, errors={errors})")

        time.sleep(SLEEP_SECONDS)

        if LIMIT and processed >= LIMIT:
            print(f"  LIMIT={LIMIT} reached, stopping")
            break

out_f.close()
print(f"\\ndone. wrote {processed} new rows to {OUTPUT_CSV} (invalid={invalid}, errors={errors})")''')


