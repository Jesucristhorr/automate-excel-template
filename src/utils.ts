import type { Sheet } from "@eyeseetea/xlsx-populate";
import type { CopyAndPasteRangeParams } from "./types";

function getLastRowNumber(worksheet: Sheet) {
    return worksheet.usedRange()?.endCell()?.rowNumber() ?? 0;
}

function copyAndPasteRange(
    {
        worksheet,
        srcRowStart,
        srcColStart,
        srcRowEnd,
        srcColEnd,
        destRowStart,
        destColStart,
    }: CopyAndPasteRangeParams
) {
    const templateRange = worksheet.range(srcRowStart, srcColStart, srcRowEnd, srcColEnd);

    // copy everything
    const templateValues = templateRange.value();
    const templateStyles = templateRange.style([
        'bold',
        'italic',
        'underline',
        'strikethrough',
        'fontSize',
        'fontFamily',
        'fontColor',
        'horizontalAlignment',
        'verticalAlignment',
        'wrapText',
        'shrinkToFit',
        'textDirection',
        'textRotation',
        'verticalText',
        'fill',
        'border',
        'numberFormat',
    ]);
    const templateRowHeights = [];
    for (let i = srcRowStart; i <= srcRowEnd; i++) {
        templateRowHeights.push(worksheet.row(i).height());
    }

    // paste everything
    const rowCount = srcRowEnd - srcRowStart + 1;
    const colCount = srcColEnd - srcColStart + 1;
    worksheet.range(destRowStart, destColStart, destRowStart + rowCount - 1, destColStart + colCount - 1)
    .value(templateValues)
    .clear() // Creates empty row
    .style(templateStyles)
    for (let i = 0; i < rowCount; i++) {
        const height = templateRowHeights[i];
        if (!height) continue;

        worksheet.row(destRowStart + i).height(height);
    }

    return;
}

export const Utils = {
    copyAndPasteRange,
    getLastRowNumber,
};
