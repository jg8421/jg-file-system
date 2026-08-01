# jg File System

Windows local file workbench for navigating folders, previewing Office documents, searching project files, and managing common file operations.

## v1.1 Highlights

- **Native PDF preview:** embedded Microsoft Edge WebView2 for smooth continuous scrolling, text selection/copy, search, and browser-quality zoom.
- **Higher-fidelity Office preview:** Word and PowerPoint use Microsoft Office's native PDF export instead of low-resolution slide images.
- **Faster switching:** the PDF engine remains alive between documents, avoiding repeated browser cold starts.
- **Long-document reading:** image-preview fallback now renders up to 16 pages at 2x resolution and keeps a 5 GB local cache for fast reopening.
- **Practical file work:** folder tabs, pinned locations, recursive search, keyboard navigation, copy/cut/paste, rename, Markdown preview/edit, and Excel cell selection/editing.

> Office preview and Excel editing require Microsoft Office installed locally. All previews stay on the local machine.

## Run

Open `Launch File Workbench.cmd` on Windows with Python 3.12 and Microsoft Office installed for Office rendering and Excel cell editing.
