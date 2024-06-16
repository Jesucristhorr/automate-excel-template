import { Utils } from '../utils';
import { NUMBER_TO_ALPHABET, CHECK_ICON_WINDINGS } from '../config/constants';
import { randomInt } from 'crypto';
import type { GroupedNewValues } from '../types';
import type { Sheet } from '@eyeseetea/xlsx-populate';

function addNewRowForTemplate(worksheet: Sheet) {
    const lastRow = Utils.getLastRowNumber(worksheet);

    Utils.copyAndPasteRange(
        {
            worksheet,
            srcRowStart: 16,
            srcRowEnd: 16,
            srcColStart: 1,
            srcColEnd: 33,
            destRowStart: lastRow,
            destColStart: 1
        }
    );

    // modify the next row so it becomes the new last row
    worksheet.row(lastRow + 1).height(30);

    return {
        rowToStart: lastRow,
    };
}

function fillData(values: GroupedNewValues, worksheet: Sheet) {
    let isFirstEntry = true;
    let count = 0;
    for (const key in values) {
        count++;
        const entries = values[key];
        if (!entries) continue;
        if (entries.length === 0) continue;

        let isFirstEntryOfGroup = true;
        let rowToStart: number;
        let currRowNumber: number;

        if (isFirstEntry) rowToStart = (Utils.getLastRowNumber(worksheet) - 1);
        else {
            const result = addNewRowForTemplate(worksheet);
            rowToStart = result.rowToStart;
        }

        currRowNumber = rowToStart;

        for (let i = 0; i < entries.length; i++) {
            const index = i as keyof typeof NUMBER_TO_ALPHABET;
            const rows: number[] = [];
            const entry = entries[i];

            if (!entry) continue;
            if (i > 0) {
                const {
                    rowToStart: currRow
                } = addNewRowForTemplate(worksheet);

                currRowNumber = currRow;
            }

            rows.push(currRowNumber);

            // create as many rows as failure modes there are
            const failureModes = entry['Modo de Falla (Causa de la Falla)'];
            const initialEffects = entry['Efecto Inicial de la Falla (Que ocurre cuando Falla)'];
            const finalEffects = entry['Efecto Final de la Falla o Consecuencia (Que ocurre cuando Falla)'];

            if (
                failureModes.length !== initialEffects.length ||
                failureModes.length !== finalEffects.length ||
                initialEffects.length !== finalEffects.length
            ) throw new Error('Failure modes are different.');

            for (let j = 0; j < failureModes.length; j++) {
                const failureMode = failureModes[j];
                const initialEffect = initialEffects[j];
                const finalEffect = finalEffects[j];

                if (j === 0) {
                    worksheet.range(currRowNumber, 5, currRowNumber, 5).value(j + 1);
                    worksheet.range(currRowNumber, 6, currRowNumber, 6).value(failureMode);
                    worksheet.range(currRowNumber, 7, currRowNumber, 7).value(j + 1);
                    worksheet.range(currRowNumber, 8, currRowNumber, 8).value(initialEffect);
                    worksheet.range(currRowNumber, 9, currRowNumber, 9).value(j + 1);
                    worksheet.range(currRowNumber, 10, currRowNumber, 10).value(finalEffect);

                    const consecuenceTypeCol = randomInt(16, 20);
                    worksheet.range(currRowNumber, consecuenceTypeCol, currRowNumber, consecuenceTypeCol).value(CHECK_ICON_WINDINGS);

                    const taskTypeCol = randomInt(20, 26);
                    worksheet.range(currRowNumber, taskTypeCol, currRowNumber, taskTypeCol).value(CHECK_ICON_WINDINGS);

                    continue;
                }

                const {
                    rowToStart: currRow
                } = addNewRowForTemplate(worksheet);

                worksheet.range(currRow, 5, currRow, 5).value(j + 1);
                worksheet.range(currRow, 6, currRow, 6).value(failureMode);
                worksheet.range(currRow, 7, currRow, 7).value(j + 1);
                worksheet.range(currRow, 8, currRow, 8).value(initialEffect);
                worksheet.range(currRow, 9, currRow, 9).value(j + 1);
                worksheet.range(currRow, 10, currRow, 10).value(finalEffect);

                const consecuenceTypeCol = randomInt(16, 20);
                worksheet.range(currRow, consecuenceTypeCol, currRow, consecuenceTypeCol).value(CHECK_ICON_WINDINGS);

                const taskTypeCol = randomInt(20, 26);
                worksheet.range(currRow, taskTypeCol, currRow, taskTypeCol).value(CHECK_ICON_WINDINGS);

                rows.push(currRow);
            }

            if (isFirstEntryOfGroup) {
                // fill function name
                const funcName = entry['Función'];
                worksheet.range(rowToStart, 1, rowToStart, 1).value(count);
                worksheet.range(rowToStart, 2, rowToStart, 2).value(funcName);
            }

            for (const row of rows) {
                // fill description
                const descriptions = entry['Descripción de las Tareas Propuestas'];
                const description = descriptions[randomInt(0, descriptions.length)];

                if (!description) throw new Error('No description.');
                worksheet.range(row, 26, row, 28).value(description).merged(true);

                // fill frequency
                const frequencies = ['7D', '2A', '1A', '6D', '6M', '1D', '2D', '3D', '1M', '2M'];
                const frequency = frequencies[randomInt(0, frequencies.length)];

                if (!frequency) throw new Error('No frequency.');
                worksheet.range(row, 29, row, 29).value(frequency);
                
                // fill responsable
                const responsibles = ['Oper.', 'Mec.', 'Pred.', 'Elec.', 'Inst.'];
                const responsible = responsibles[randomInt(0, responsibles.length)];

                if (!responsible) throw new Error('No responsible.');
                worksheet.range(row, 30, row, 30).value(responsible);

                // fill respQty
                const respQty = randomInt(1, 4);
                worksheet.range(row, 31, row, 31).value(respQty);

                // fill men hours
                const menHours = respQty > 1 ? '2x' + String(randomInt(2, 5)) : '1x1';
                worksheet.range(row, 32, row, 32).value(menHours);

                // fill op team
                const opTeam = ['Mec.', 'Oper.'].includes(responsible) ? 'Sí' : 'No';
                worksheet.range(row, 33, row, 33).value(opTeam);
            }

            const firstRow = rows.at(0);
            const finalRow = rows.at(-1);

            if (!firstRow || !finalRow) throw new Error('No rows.');

            // merge functional failure
            const functionalFailure = entry['Falla Funcional'];
            worksheet.range(firstRow, 4, finalRow, 4).value(functionalFailure).merged(true);

            // fill functional failure
            const alphabetLetter = NUMBER_TO_ALPHABET[index] ?? '---';
            worksheet.range(firstRow, 3, finalRow, 3).value(alphabetLetter).merged(true);

            currRowNumber = finalRow;

            isFirstEntryOfGroup = false;
        }

        const funcNameIndexRange = worksheet.range(rowToStart, 1, currRowNumber, 1);
        const funcNameRange = worksheet.range(rowToStart, 2, currRowNumber, 2);
        const funcNameIndex = funcNameIndexRange.value();
        const funcName = funcNameRange.value();
        funcNameRange.value(funcName).style({ bold: true }).merged(true);
        funcNameIndexRange.value(funcNameIndex).merged(true);

        isFirstEntry = false;
    }

    return;
}

export default {
    addNewRowForTemplate,
    fillData,
};
