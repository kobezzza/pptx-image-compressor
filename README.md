# pptx-image-compressor

A tiny CLI utility to compress images embedded in PowerPoint (`.pptx/.pptm`) files. It unpacks the PPTX, resizes and recompresses `.jpg/.jpeg/.png` images.

- Supports JPEG and PNG compression
- Resizes images to a maximum dimension while keeping an aspect ratio
- Adjustable quality/compression levels
- Creates a new compressed file, preserving the original
- Shows detailed compression statistics

## Requirements

- Node.js 18+
- One of the following for ZIP operations:
  - macOS/Linux: `unzip` and `zip` installed (usually preinstalled on macOS and many Linux distros), OR
  - Windows: PowerShell 5+ (default on modern Windows) — the tool will use `Expand-Archive`/`Compress-Archive` automatically

## Installation

```
npm install pptx-image-compressor
```

## Usage

```
pptx-image-compressor <path_to_file.pptx> [max_size] [quality]
```

Parameters:

- path_to_file.pptx: Path to the PowerPoint file you want to compress
- max_size: Maximum width/height in pixels for images (default: 1920)
- quality:
  - JPEG: quality 1–100 (higher is better quality/larger size). Default: 80
  - PNG: inverse scale where 100 ≈ the lowest compression and 0 ≈ the highest compression; the tool maps this to PNG compression levels 0–9

Examples:

- Keep images up to 1920 px on the long side at default quality:
  ```
  node main.js "My Presentation.pptx"
  ```

- Resize to 1600 px and use JPEG quality 75 (and corresponding PNG compression):
  ```
  node main.js ".\slides\deck.pptx" 1600 75
  ```

On success, you’ll see a summary like:

```
✅ Done! File saved as: <original_name>_compressed.pptx
📊 Compression results:
   Original size: <X> MB
   New size:      <Y> MB
   Saved:         <Z> MB
   Compression ratio: <R>%
```

The tool creates a new file next to the original, suffixed with `_compressed` (e.g., `deck_compressed.pptx`). Your original file is left untouched.

## License

MIT — see LICENSE.
