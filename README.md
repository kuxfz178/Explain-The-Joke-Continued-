# Re-evaluating Humour Understanding in Vision–Language Models

This repository contains the code, annotations, and model outputs behind our bachelor thesis at the University of Copenhagen, *Re-evaluating Humour Understanding in Vision–Language Models: A Balanced Human-Annotated 501-Image MCQ Extension of Alm (2026)*.

The thesis extends Frida Alm's *ExplainTheJoke* benchmark to a 501-item, human-annotated multiple-choice setting with balanced humour-style and contextual-knowledge coverage. Seven vision–language models are evaluated, and the analysis combines overall accuracy, a literal-trap diagnostic, category-level breakdowns using primary and secondary labels, and pairwise paid-API comparisons.

## What is in here

### Pipeline scripts

* `build.py` reads the annotation spreadsheet, repairs encoding artefacts, strips wrapping quotes from every text field, and writes the three benchmark CSVs. The correct-answer letter is assigned with a balanced, seeded shuffle so that A, B, C, and D are distributed evenly across the 501 items.
* `merge_secondary.py` overlays the secondary humour and context annotations from the updated source spreadsheet onto `answer_key.csv` without touching the existing correct-letter assignments. We used this when the secondary labels were added in a second annotation pass.
* `build_openrouter_notebook.py` is the generator script for the Colab notebook below. We kept it in the repo so that the notebook can be rebuilt without manual editing.
* `run_mcq_openrouter.ipynb` is the Colab notebook that calls every model through the OpenRouter API. It is resume-safe, so if a run is interrupted you can just rerun the loop and it picks up where it left off.
* `grade.py` joins each model's answer CSV against `answer_key.csv` and writes a multi-sheet Excel report covering the leaderboard, breakdowns by humour style and contextual knowledge (primary and secondary), per-item details, and difficulty rankings.
* `stats.py` computes the bootstrap confidence intervals, pairwise McNemar tests, and the methodology sheet that documents how each number is calculated.

### Data

* `source.xlsx` is the cleaned annotation spreadsheet with the correct explanation, three distractors, literal description, and humour-style and contextual-knowledge labels for every item.
* `annotations_clean.csv`, `model_input.csv`, and `answer_key.csv` are the three files produced by `build.py`. `model_input.csv` is the only file the models see.
* `metadata.csv` maps each File id to its image path, which the Colab notebook uses to load the right image for each prompt.
* `runs/` contains one CSV per model with the parsed answer letter for each of the 501 items.

### Reports

* The Excel reports produced by `grade.py` and `stats.py` are included so that the numbers in Chapter 5 can be verified without rerunning anything.

## How to reproduce

1. Install the Python dependencies with `pip install -r requirements.txt`.
2. Sign up at [openrouter.ai](https://openrouter.ai), add a few dollars of credit, and create an API key.
3. Open `run_mcq_openrouter.ipynb` in Colab. Paste your API key into Colab Secrets as `OPENROUTER_API_KEY` and point `PROJECT_DIR` at your copy of this folder on Google Drive.
4. Run `build.py` once to regenerate the benchmark CSVs from `source.xlsx`.
5. Run the notebook seven times, changing the `MODEL` string in the config cell each time. Each run writes its output CSV into `runs/`.
6. Run `grade.py` and `stats.py` locally. The Excel reports drop into the same folder.

The full pipeline runs end-to-end in well under an hour on a normal laptop, since all inference happens server-side at OpenRouter and no GPU is needed.

## Models evaluated

All seven models were called on 2 June 2026 through OpenRouter using a single OpenAI-compatible endpoint. The OpenRouter IDs are:

* `openai/gpt-5`
* `anthropic/claude-opus-4-6`
* `google/gemini-3.5-flash`
* `google/gemma-3-4b-it`
* `google/gemma-4-26b-a4b-it`
* `qwen/qwen3-vl-30b-a3b-thinking`
* `qwen/qwen3.5-9b`

Temperature was fixed at 0. `max_tokens` was set to 2000 to give the reasoning-capable models enough room for hidden thinking tokens without blowing up the cost.

## A few notes on what is not in the repo

The 501 meme images are not redistributed here. They originate from Frida Alm's *ExplainTheJoke* dataset, and the right way to access them is through her release rather than ours. `metadata.csv` documents the File ids we used so that the images can be matched back to her dataset.

API keys are deliberately not committed. Put your OpenRouter key in Colab Secrets, not in a file.

## Authors

Naomi Jenfort and Sabrina Ismail, BSc in Computer Science and Economics, University of Copenhagen, June 2026. Supervisor: Desmond Elliott.
