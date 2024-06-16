import { VALUES_FILE_FULL_PATH } from "../config/constants";
import { GroupedNewValues, NewValue, PromiseResult } from "../types";

import XLSX from '@e965/xlsx';

function parseLists(values: NewValue[]) {
    return values.map((value: any) => {
        const propertiesToParse = [
            'Modo de Falla (Causa de la Falla)',
            'Efecto Inicial de la Falla (Que ocurre cuando Falla)',
            'Efecto Final de la Falla o Consecuencia (Que ocurre cuando Falla)',
        ] as (keyof NewValue)[];

        // This is for the list type
        for (const property of propertiesToParse) {
            const rawValue = value[property];

            if (typeof rawValue !== 'string') continue;
            
            const regex = /\d+\.\s([^0-9]+)/g;
            let match;
            const results = [];
    
            while ((match = regex.exec(rawValue)) !== null) {
                const result = match[1]?.trim();
                const lastChar = result?.substring(result.length - 1);
                if (!result) continue;
                results.push(lastChar === '.' ? result : result + '.');
            }

            value[property] = results;
        }

        // This is for the comma type of lists
        const rawValue: string = value['Descripción de las Tareas Propuestas'] ?? '';
        const regex = /,| y /g;
        const opciones = rawValue.split(regex).map(tasks => tasks.trim());

        // Capitalize
        value['Descripción de las Tareas Propuestas'] = opciones.map(tasks => tasks.charAt(0).toUpperCase() + tasks.slice(1));

        return value;
    });
}

function groupValuesByFunction(values: NewValue[]): GroupedNewValues {
    return values.reduce<GroupedNewValues>((prevValue, currVal) => {
        const funcName = currVal['Función'];

        if (!prevValue[funcName]) prevValue[funcName] = [];

        prevValue[funcName]?.push(currVal);

        return prevValue;
    }, {});
}

async function getValuesToFill(): PromiseResult<GroupedNewValues> {
    try {
        const workbook = XLSX.readFile(VALUES_FILE_FULL_PATH);

        const worksheet = workbook.Sheets['vall'];

        if (!worksheet) throw new Error('No worksheet in values file.');

        const data = XLSX.utils.sheet_to_json(worksheet) as NewValue[];

        const parsedValues = parseLists(data) as unknown as NewValue[];
        
        return [null, groupValuesByFunction(parsedValues)];
    } catch (error) {
        return [error, null];
    }
}

export default {
    getValuesToFill,
};
