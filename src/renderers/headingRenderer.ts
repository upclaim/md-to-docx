import { Paragraph, HeadingLevel, Bookmark, AlignmentType } from "docx";
import { Style, HeadingConfig } from "../types.js";
import { sanitizeForBookmarkId } from "../utils/bookmarkUtils.js";
import { processFormattedTextForHeading } from "../parsers/textParser.js";

/**
 * Processes a heading line and returns appropriate paragraph formatting and a bookmark ID
 * @param line - The heading line to process
 * @param config - The heading configuration
 * @param style - The style configuration
 * @param documentType - The document type
 * @returns An object containing the processed paragraph and its bookmark ID
 */
export function processHeading(
	line: string,
	config: HeadingConfig,
	style: Style,
	documentType: "document" | "report",
): { paragraph: Paragraph; bookmarkId: string } {
	const headingText = line.replace(new RegExp(`^#{${config.level}} `), "");
	const headingLevel = config.level;
	// Generate a unique bookmark ID using the clean text (without markdown)
	const cleanTextForBookmark = headingText
		.replace(/\*\*/g, "")
		.replace(/\*/g, "");
	const bookmarkId = `_Toc_${sanitizeForBookmarkId(
		cleanTextForBookmark,
	)}_${Date.now()}`;

	// Get the appropriate font size based on heading level and custom style
	let headingSize = style.titleSize;

	// Use specific heading size if provided, otherwise calculate based on level
	if (headingLevel === 1 && style.heading1Size) {
		headingSize = style.heading1Size;
	} else if (headingLevel === 2 && style.heading2Size) {
		headingSize = style.heading2Size;
	} else if (headingLevel === 3 && style.heading3Size) {
		headingSize = style.heading3Size;
	} else if (headingLevel === 4 && style.heading4Size) {
		headingSize = style.heading4Size;
	} else if (headingLevel === 5 && style.heading5Size) {
		headingSize = style.heading5Size;
	} else if (headingLevel > 1) {
		// Fallback calculation if specific size not provided
		headingSize = style.titleSize - (headingLevel - 1) * 4;
	}

	// Determine alignment based on heading level
	let alignment;

	// Check for level-specific alignment first
	if (headingLevel === 1 && style.heading1Alignment) {
		alignment =
			AlignmentType[style.heading1Alignment as keyof typeof AlignmentType];
	} else if (headingLevel === 2 && style.heading2Alignment) {
		alignment =
			AlignmentType[style.heading2Alignment as keyof typeof AlignmentType];
	} else if (headingLevel === 3 && style.heading3Alignment) {
		alignment =
			AlignmentType[style.heading3Alignment as keyof typeof AlignmentType];
	} else if (headingLevel === 4 && style.heading4Alignment) {
		alignment =
			AlignmentType[style.heading4Alignment as keyof typeof AlignmentType];
	} else if (headingLevel === 5 && style.heading5Alignment) {
		alignment =
			AlignmentType[style.heading5Alignment as keyof typeof AlignmentType];
	} else if (style.headingAlignment) {
		// Fallback to general heading alignment if no level-specific alignment
		alignment =
			AlignmentType[style.headingAlignment as keyof typeof AlignmentType];
	}

	// Process the heading text to handle markdown formatting (bold/italic)
	const processedTextRuns = processFormattedTextForHeading(
		headingText,
		headingSize,
		style,
	);

	// Create the paragraph with bookmark
	const paragraph = new Paragraph({
		children: [
			new Bookmark({
				id: bookmarkId,
				children: processedTextRuns,
			}),
		],
		heading:
			headingLevel as unknown as (typeof HeadingLevel)[keyof typeof HeadingLevel],
		spacing: {
			before:
				config.level === 1 ? style.headingSpacing * 2 : style.headingSpacing,
			after: style.headingSpacing / 2,
		},
		alignment: alignment,
		style: `Heading${headingLevel}`, // This is crucial for TOC recognition
		bidirectional: style.direction === "RTL",
	});

	return { paragraph, bookmarkId };
}
