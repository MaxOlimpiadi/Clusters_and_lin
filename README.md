# Agentivity-Oriented Data Transformation for German Literary Texts

This repository contains a data processing pipeline that transforms manually annotated JSON data into a structured JSON format tailored for downstream experiments on **agentivity detection in German literary texts**.

The project focuses on converting human-annotated character-level information into a normalized representation that aligns with the internal data schema used in the agentivity analysis framework.

---

## Data Description

- **Input**: JSON files containing manual annotations of literary texts, including character mentions, participation objects and agentivity-related labels.
- **Annotations**: Created manually by domain annotators.
- **Language**: German
- **Domain**: Narrative and literary prose

---


## Output Format

The new JSON files follow a project-specific schema designed to support agentivity classification using seq2seq models.
