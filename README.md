# PowerShell DOCX to Markdown Converter

A sophisticated PowerShell script that uses Pandoc to convert structured Microsoft Word documents into a nested, multi-file Markdown project, perfect for personal wikis, static site generators, or Obsidian vaults.

This script is designed for writers and developers who manage large, structured documents in Microsoft Word and need an automated way to convert them into a clean, chapter-based Markdown format.

---

## Features

* **Automated Conversion:** Processes a single `.docx` file or an entire directory of them in one command.
* **Structured Output:** Automatically creates a nested folder structure based on Heading 1 (Book), Heading 2 (Chapter), and Heading 3 (Part) styles.
* **YAML Frontmatter Support:** Intelligently parses YAML frontmatter from the source Word document and preserves it in the final Markdown files.
* **Automatic Navigation:** Generates `previous` and `next` navigation links in the frontmatter of each chapter, with an option for Obsidian-friendly wiki-links.
* **Robust & Configurable:** Includes command-line parameters for configuration (`-Padding`, `-PurgeOutputFolder`) and a full, `man` page-style help screen.

## Prerequisites

This script requires **Pandoc** to be installed on your system and accessible in your environment's PATH. You can download it from [pandoc.org](https://pandoc.org/).

## Usage

The script is designed to be run from a PowerShell terminal in the directory containing your `.docx` files.

### Parameters

* `-DocxFile <string>`
    * (Optional) The name of a specific Word document to process. If omitted, all `.docx` files in the current directory are processed.
* `-Padding <int>`
    * (Optional) The number of leading zeros for folder/file prefixes. Defaults to `3`.
* `-PurgeOutputFolder`
    * (Optional) If present, deletes all files/folders in the output directory before creating new ones.
* `-ObsidianFriendlyLinks`
    * (Optional) If present, creates simple `[[filename]]` links that rely on Obsidian's internal linking.
* `-Help`, `-?`
    * (Optional) Displays the help message and exits.

### Examples

**Process a single file with default settings:**
```powershell
.\Convert-And-Split.ps1 -DocxFile 'MyNovel.docx'