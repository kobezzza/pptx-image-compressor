#!/usr/bin/env node

import fs from "fs/promises";

import path from "path";

import { execSync } from "child_process";

import sharp from "sharp";

const [inputPath, ...args] = process.argv.slice(2);

const maxDimension = parseInt(args[0] ?? "1920");

const quality = parseInt(args[1] ?? "80");

if (!inputPath) {
	console.error(`
Usage: node compress-pptx-images.js <path_to_file.pptx> [max_size] [quality]

  path_to_file.pptx : Path to your PowerPoint file
  max_size          : Maximum width/height in pixels (default: 1920)
  quality           : Quality for JPEG (1–100) and PNG (100–0, where 0 is max compression)

Example: pptx-image-compressor "My Presentation.pptx" 1920 80
`);

	process.exit(1);
}

await compressPptxImages();

async function compressPptxImages() {
	const tempDir = path.join(process.cwd(), `temp_pptx_${Date.now()}`);

	try {
		if (!await fileExists(inputPath)) {
			console.error(`Error: File "${inputPath}" not found!`);
			process.exit(1);
		}

		const inputExtname = path.extname(inputPath);

		const outputFile = path.join(
			path.dirname(inputPath),
			`${path.basename(inputPath, inputExtname)}_compressed${inputExtname}`
		);

		console.log("🔄 Creating temporary directory...");
		await fs.mkdir(tempDir, { recursive: true });

		console.log("📦 Unpacking PPTX as ZIP...");
		const tempZip = path.join(tempDir, "presentation.zip");
		await fs.copyFile(inputPath, tempZip);

		const extractDir = path.join(tempDir, "extracted");
		await fs.mkdir(extractDir);

		// Cross-platform extraction
		try {
			execSync(`unzip -q "${tempZip}" -d "${extractDir}"`);

		} catch {
			// For Windows, we use PowerShell
			execSync(`Expand-Archive -Path "${tempZip}" -DestinationPath "${extractDir}" -Force`, {shell: "powershell.exe"});
		}

		console.log("🔍 Searching and compressing images...");
		const mediaDir = path.join(extractDir, "ppt", "media");
		await processImagesInDirectory(mediaDir, maxDimension, quality);

		const slidesMediaDir = path.join(extractDir, "ppt", "slides", "_media");

		if (await fileExists(slidesMediaDir)) {
			await processImagesInDirectory(slidesMediaDir, maxDimension, quality);
		}

		console.log("📦 Repacking to PPTX...");
		const cwd = process.cwd();
		process.chdir(extractDir);

		// Cross-platform packing
		try {
			execSync(`zip -qr "${tempZip}" *`);

		} catch {
			// For Windows, we use PowerShell
			const files = await fs.readdir(".");
			execSync(`Compress-Archive -LiteralPath ${files.map(f => `"${f}"`).join(",")} -DestinationPath "${tempZip}" -Force`, {shell: "powershell.exe"});
		}

		process.chdir(cwd);
		await fs.copyFile(tempZip, outputFile);

		const originalStats = await fs.stat(inputPath);
		const newStats = await fs.stat(outputFile);
		const saved = (originalStats.size - newStats.size) / 1024 / 1024;

		console.log(`
✅ Done! File saved as: ${outputFile}

📊 Compression results:
   Original size: ${(originalStats.size / 1024 / 1024).toFixed(2)} MB
   New size:      ${(newStats.size / 1024 / 1024).toFixed(2)} MB
   Saved:         ${saved.toFixed(2)} MB
   Compression ratio: ${((1 - newStats.size / originalStats.size) * 100).toFixed(1)}%
`);

	} catch (error) {
		console.error("❌ An error occurred:", error.message);
		process.exit(1);

	} finally {
		try {
			console.log("🧹 Cleaning up temporary files...");
			await fs.rm(tempDir, {recursive: true, force: true});
		} catch {}
	}
}

async function processImagesInDirectory(dir, maxDimension, quality) {
	try {
		const files = await fs.readdir(dir);

		for (const file of files) {
			const filePath = path.join(dir, file);
			const stats = await fs.stat(filePath);

			if (!stats.isFile()) {
				continue;
			}

			const ext = path.extname(file).toLowerCase();

			if (![".jpg", ".jpeg", ".png"].includes(ext)) {
				continue;
			}

			try {
				console.log(`Processing: ${file} (${(stats.size / 1024 / 1024).toFixed(2)} MB)`);

				const image = sharp(filePath);
				const metadata = await image.metadata();

				if (metadata.width <= maxDimension && metadata.height <= maxDimension) {
					continue;
				}

				const resizeOptions = {
					width: metadata.width > metadata.height ? maxDimension : null,
					height: metadata.height > metadata.width ? maxDimension : null,
					fit: "inside",
					withoutEnlargement: true
				};

				const tempPath = `${filePath}.temp`;

				if (ext === ".png") {
					const compressionLevel = Math.floor((100 - quality) / 11.1); // Convert 0-100 to 0-9

					await image
						.resize(resizeOptions)

						.png({
							compressionLevel: Math.max(0, Math.min(9, compressionLevel)),
							adaptiveFiltering: true,
							force: false
						})

						.toFile(tempPath);

				} else {
					await image
						.resize(resizeOptions)

						.jpeg({
							quality: quality,
							mozjpeg: true
						})

						.toFile( tempPath);
				}

				const newStats = await fs.stat(tempPath);

				if (stats.size - newStats.size >= 0) {
					await fs.unlink(filePath);
					await fs.rename(tempPath, filePath);

					const saved = (stats.size - newStats.size) / 1024 / 1024;
					console.log(`✅ Compressed: ${(stats.size / 1024 / 1024).toFixed(2)} MB → ${(newStats.size / 1024 / 1024).toFixed(2)} MB (saved: ${saved.toFixed(2)} MB)`);

				} else {
					await fs.unlink(tempPath);
				}

			} catch (imageError) {
				console.warn(`⚠️ Failed to process ${file}:`, imageError.message);
			}
		}

	} catch (dirError) {
		console.warn(`⚠️ Failed to process directory ${dir}:`, dirError.message);
	}
}

async function fileExists(filePath) {
	try {
		await fs.access(filePath);
		return true;

	} catch {
		return false;
	}
}
