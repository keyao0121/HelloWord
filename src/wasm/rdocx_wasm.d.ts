/* tslint:disable */
/* eslint-disable */

/**
 * A Word document (.docx) that can be created, modified, and exported.
 */
export class WasmDocument {
    free(): void;
    [Symbol.dispose](): void;
    /**
     * Add a paragraph with bold text.
     */
    addBoldParagraph(text: string): void;
    /**
     * Add a heading paragraph (level 1-6).
     */
    addHeading(text: string, level: number): void;
    /**
     * Add a paragraph with the given text.
     */
    addParagraph(text: string): void;
    /**
     * Add a simple table with the given number of rows and columns.
     */
    addTable(rows: number, cols: number): void;
    /**
     * Mark the start of one undoable edit. Body edits made until `commitEdit` form one step.
     */
    beginEdit(): void;
    canRedo(): boolean;
    canUndo(): boolean;
    /**
     * Close the edit `beginEdit` opened. Returns 0 when the body did not change, 1 when a
     * new undo step was recorded, and 2 when `merge` folded the change into the previous
     * step, which is how a run of typed characters becomes one step.
     */
    commitEdit(merge: boolean): number;
    /**
     * Delete the inline picture `index` of a paragraph. The paragraph stays.
     */
    deleteImage(para: number, index: number): void;
    /**
     * Delete the characters `[start, end)` of a paragraph's text.
     */
    deleteRange(para: number, start: number, end: number): void;
    /**
     * Delete the column of the cell of a paragraph in a table, or the whole table when
     * that is its only column, and return the index of a paragraph where the cell was.
     */
    deleteTableColumn(para: number): number;
    /**
     * Delete the row of a paragraph in a table, or the whole table when that is its only
     * row, and return the index of a paragraph where the row was.
     */
    deleteTableRow(para: number): number;
    /**
     * Return the file bytes of one font of the current layout.
     */
    fontData(id: number): Uint8Array;
    /**
     * Return the fonts the current layout draws with, as a JSON array.
     */
    fontsJson(): string;
    /**
     * Return what a toolbar shows for a position in a paragraph, as JSON: whether the
     * character at the offset is underlined, the paragraph's own alignment (or null when
     * it inherits one) and its heading level (0 when it is not a heading).
     */
    formatJson(para: number, offset: number): string;
    /**
     * Set one character property on the characters `[start, end)` of a paragraph.
     * `property` is `bold`, `italic` or `underline` with the value `1` or `0`, `size` with
     * a size in points, `color` with six hex digits or `auto`, `highlight` with a Word
     * highlight colour name or `none`, or `font` with a family.
     */
    formatRange(para: number, start: number, end: number, property: string, value: string): void;
    /**
     * Open a document from DOCX bytes.
     */
    static fromBytes(data: Uint8Array): WasmDocument;
    /**
     * Get the text content of the entire document.
     */
    getText(): string;
    /**
     * Return the inline pictures of a paragraph in document order, as a JSON array of
     * `[width, height]` in points.
     */
    imagesJson(para: number): string;
    /**
     * Insert a picture in a paragraph of its own, after the body child that holds a
     * paragraph. The size is in points.
     */
    insertImage(para: number, data: Uint8Array, filename: string, width: number, height: number): void;
    /**
     * Insert a table with single-line borders and cell padding after the body child that holds a paragraph,
     * and keep a paragraph after it, which a document needs after a table.
     */
    insertTable(para: number, rows: number, columns: number): void;
    /**
     * Insert an empty column left or right of the cell of a paragraph in a table, and
     * return the index of the first paragraph of the new cell in the same row.
     */
    insertTableColumn(para: number, after: boolean): number;
    /**
     * Insert an empty row above or below the row of a paragraph in a table, formatted as
     * that row, and return the index of the first paragraph of the new row's cell in the
     * same column.
     */
    insertTableRow(para: number, below: boolean): number;
    /**
     * Insert text into a paragraph at a character offset of its text.
     */
    insertText(para: number, offset: number, text: string): void;
    /**
     * Return the bytes of one image drawn on a page, by the id `pageJson` reports.
     */
    mediaData(page: number, id: string): Uint8Array;
    /**
     * Join the next sibling of a paragraph onto its end and return the character offset
     * of the join. The first paragraph keeps its properties.
     */
    mergeParagraphs(para: number): number;
    /**
     * Merge the cells of a table between the cells of two paragraphs into one, keeping
     * what they hold, and return the index of the first paragraph of the merged cell.
     */
    mergeTableCells(from: number, to: number): number;
    /**
     * Move the inline picture `index` of a paragraph into a paragraph of its own, placed
     * before or after the body child that holds paragraph `target`. A paragraph the
     * picture leaves empty goes with it. Both paragraphs are direct body paragraphs or
     * in tables of the body; the picture always lands in the body itself. Returns the
     * index of the paragraph the picture is in afterwards.
     */
    moveImage(para: number, index: number, target: number, after: boolean): number;
    /**
     * Create a new, empty document.
     */
    constructor();
    /**
     * Lay the document out and return its page count.
     */
    pageCount(): number;
    /**
     * Return one laid-out page as JSON: its size in points and its positioned elements.
     */
    pageJson(index: number): string;
    /**
     * Return every page's size in points, as `[width0, height0, width1, height1, ...]`.
     */
    pageSizes(): Float64Array;
    /**
     * Get the number of paragraphs in the document.
     */
    paragraphCount(): number;
    /**
     * Return the text of a paragraph, in the characters its offsets count.
     */
    paragraphText(para: number): string;
    /**
     * Apply again the edit `undo` took back. Returns false when there is none.
     */
    redo(): boolean;
    /**
     * Remove a paragraph, such as one inside a deleted selection. A table cell keeps its
     * last paragraph, which the file format requires.
     */
    removeParagraph(para: number): void;
    /**
     * Replace all occurrences of a placeholder with a value.
     */
    replacePlaceholder(placeholder: string, value: string): number;
    /**
     * Replace every occurrence of a text in the document and return how many there were.
     */
    replaceText(find: string, replacement: string): number;
    /**
     * Set the size, in points, of the inline picture `index` of a paragraph.
     */
    resizeImage(para: number, index: number, width: number, height: number): void;
    /**
     * Set a paragraph's alignment: `left`, `center`, `right` or `justify`.
     */
    setAlignment(para: number, alignment: string): void;
    /**
     * Set how far a paragraph's first line starts in from the rest, in points; 0 removes it.
     */
    setFirstLineIndent(para: number, points: number): void;
    /**
     * Give a paragraph the document's heading style of a level from 1, or its default
     * paragraph style for level 0.
     */
    setHeading(para: number, level: number): void;
    /**
     * Set a paragraph's left indent in points; 0 removes it.
     */
    setIndent(para: number, points: number): void;
    /**
     * Set a paragraph's line spacing as a multiple of single spacing.
     */
    setLineSpacing(para: number, multiple: number): void;
    /**
     * Make a paragraph an item of a bullet list or a numbered list, or take it out of its
     * list: `bullet`, `number` or `none`. An item right after an item of the same kind
     * joins that list; a numbered item anywhere else starts a list that counts from 1.
     */
    setList(para: number, kind: string): void;
    /**
     * Start a paragraph on a new page, or stop doing so.
     */
    setPageBreakBefore(para: number, on: boolean): void;
    /**
     * Set the space before and after a paragraph, in points. A negative number leaves
     * that side to the paragraph's style.
     */
    setParagraphSpacing(para: number, before: number, after: number): void;
    /**
     * Split a paragraph in two at a character offset. The second half becomes its next
     * sibling, and so paragraph `para + 1`, with the same paragraph properties.
     */
    splitParagraph(para: number, offset: number): void;
    /**
     * Split the merged cell of a paragraph in a table back into the cells it covers, and
     * return the index of the first paragraph of the cell that is left in its place.
     */
    splitTableCell(para: number): number;
    /**
     * Return what an editor needs to map the layout back to the document, as JSON.
     *
     * `paras` lists every paragraph an editor can put a caret in, in document order: the
     * direct body paragraphs and the paragraphs in table cells. Each has its length in
     * characters (`len`), its heading level or 0 (`h`), whether the next entry is its next
     * sibling and so can be joined onto it (`m`), whether it is in a table (`t`), whether
     * it is a list item (`l`), and the
     * page numbers, from 1, that the body child it belongs to occupies (`pages`).
     * The index of a paragraph in `paras` is what every editing call takes.
     * `nodes[i]` is the `paras` index of source node `i + 1`, or -1 for other text.
     */
    structureJson(): string;
    /**
     * Export as DOCX bytes.
     */
    toDocxBytes(): Uint8Array;
    /**
     * Convert to a complete HTML document string.
     */
    toHtml(): string;
    /**
     * Convert to an HTML fragment (body content only).
     */
    toHtmlFragment(): string;
    /**
     * Convert to Markdown.
     */
    toMarkdown(): string;
    /**
     * Render the document as PDF bytes.
     */
    toPdf(): Uint8Array;
    /**
     * Take back the last recorded edit. Returns false when there is none.
     */
    undo(): boolean;
}

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_wasmdocument_free: (a: number, b: number) => void;
    readonly wasmdocument_addBoldParagraph: (a: number, b: number, c: number) => void;
    readonly wasmdocument_addHeading: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_addParagraph: (a: number, b: number, c: number) => void;
    readonly wasmdocument_addTable: (a: number, b: number, c: number) => void;
    readonly wasmdocument_beginEdit: (a: number) => void;
    readonly wasmdocument_canRedo: (a: number) => number;
    readonly wasmdocument_canUndo: (a: number) => number;
    readonly wasmdocument_commitEdit: (a: number, b: number) => number;
    readonly wasmdocument_deleteImage: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_deleteRange: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly wasmdocument_deleteTableColumn: (a: number, b: number, c: number) => void;
    readonly wasmdocument_deleteTableRow: (a: number, b: number, c: number) => void;
    readonly wasmdocument_fontData: (a: number, b: number, c: number) => void;
    readonly wasmdocument_fontsJson: (a: number, b: number) => void;
    readonly wasmdocument_formatJson: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_formatRange: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => void;
    readonly wasmdocument_fromBytes: (a: number, b: number, c: number) => void;
    readonly wasmdocument_getText: (a: number, b: number) => void;
    readonly wasmdocument_imagesJson: (a: number, b: number, c: number) => void;
    readonly wasmdocument_insertImage: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => void;
    readonly wasmdocument_insertTable: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly wasmdocument_insertTableColumn: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_insertTableRow: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_insertText: (a: number, b: number, c: number, d: number, e: number, f: number) => void;
    readonly wasmdocument_mediaData: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly wasmdocument_mergeParagraphs: (a: number, b: number, c: number) => void;
    readonly wasmdocument_mergeTableCells: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_moveImage: (a: number, b: number, c: number, d: number, e: number, f: number) => void;
    readonly wasmdocument_new: () => number;
    readonly wasmdocument_pageCount: (a: number, b: number) => void;
    readonly wasmdocument_pageJson: (a: number, b: number, c: number) => void;
    readonly wasmdocument_pageSizes: (a: number, b: number) => void;
    readonly wasmdocument_paragraphCount: (a: number) => number;
    readonly wasmdocument_paragraphText: (a: number, b: number, c: number) => void;
    readonly wasmdocument_redo: (a: number) => number;
    readonly wasmdocument_removeParagraph: (a: number, b: number, c: number) => void;
    readonly wasmdocument_replacePlaceholder: (a: number, b: number, c: number, d: number, e: number) => number;
    readonly wasmdocument_replaceText: (a: number, b: number, c: number, d: number, e: number) => number;
    readonly wasmdocument_resizeImage: (a: number, b: number, c: number, d: number, e: number, f: number) => void;
    readonly wasmdocument_setAlignment: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly wasmdocument_setFirstLineIndent: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_setHeading: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_setIndent: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_setLineSpacing: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_setList: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly wasmdocument_setPageBreakBefore: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_setParagraphSpacing: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly wasmdocument_splitParagraph: (a: number, b: number, c: number, d: number) => void;
    readonly wasmdocument_splitTableCell: (a: number, b: number, c: number) => void;
    readonly wasmdocument_structureJson: (a: number, b: number) => void;
    readonly wasmdocument_toDocxBytes: (a: number, b: number) => void;
    readonly wasmdocument_toHtml: (a: number, b: number) => void;
    readonly wasmdocument_toHtmlFragment: (a: number, b: number) => void;
    readonly wasmdocument_toMarkdown: (a: number, b: number) => void;
    readonly wasmdocument_toPdf: (a: number, b: number) => void;
    readonly wasmdocument_undo: (a: number) => number;
    readonly __wbindgen_export: (a: number, b: number) => number;
    readonly __wbindgen_export2: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_add_to_stack_pointer: (a: number) => number;
    readonly __wbindgen_export3: (a: number, b: number, c: number) => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
