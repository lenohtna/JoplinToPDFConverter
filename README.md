# Joplin to PDF Converter

## A tool that converts Joplin export files (\*.jex) into PDF format

### Background

I originally created this script out of necessity. As a longtime Samsung Galaxy Note user, I stored all my notes in **S Note** (now **Samsung Notes**). However, after my Galaxy Note 10+ started slowing down due to multiple drops, I temporarily switched to a more affordable device—the **Xiaomi Redmi Note 12 Turbo**. Since Xiaomi doesn’t support Samsung Notes, I had to find a cross-platform note-taking app that worked on both mobile and PC. That’s when I discovered **Joplin**, which became my go-to app for over a year. I even got to learn its **Markdown formatting** along the way.

In February 2025, I upgraded to a **Samsung Galaxy S24 Ultra**, finally regaining access to Samsung Notes and my old notes. However, I faced a new challenge: **how to transfer my Joplin notes to Samsung Notes**.

After extensive research, I found a utility to convert Markdown files to PDF. While this sounded promising, Joplin’s export files include tags and structured folders. Because Samsung Notes can import multiple PDFs in one go, I needed a way to **convert all Markdown files at once while maintaining my folder structure**—and that’s why I developed this script.

### Why Share It?

I may be the **only one dealing with this situation**, but if someone else happens to face the same challenge, I hope this script can help. I developed this script based on my current use case, but you're free to explore how I coded it and modify it according to your own requirements.

### Features

This tool bulk-converts Markdown files from a **Joplin export file** into **PDFs**, organizing them into folders based on their respective **Notebooks**.

### Requirements

Since this is a **VBScript**, you’ll need a **Windows PC**. The following software must be installed:

1. **7zip** – Used to extract the Joplin export file (\*.jex) as an archive.
2. **Pandoc** – The actual file converter. Once installed, it can be run via **Windows Terminal** or **Command Line**. You can install it using UniGetUI or with this command:

`winget.exe install --id "JohnMacFarlane.Pandoc" --exact --source winget --accept-source-agreements --disable-interactivity --silent --accept-package-agreements –force`

3. **.NET 7.0 or higher** – Required for the script to function properly, as Pandoc utilizes the .NET Framework.

### How to Use

1. **Download the zip file** and extract its contents. Ensure that **JoplinToPDFConverter.vbs** and **MarkdownToPDF.bat** are in the same folder.
    - **JoplinToPDFConverter.vbs** is the main script.
    - **MarkdownToPDF.bat** is used by the script for conversion.
2. **Run JoplinToPDFConverter.vbs**, and the process will begin.
