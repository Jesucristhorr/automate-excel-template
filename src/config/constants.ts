import path from 'path';

export const CHECK_ICON_WINDINGS = 'l';
export const VALUES_FILENAME = 'values.xlsx';
export const TEMPLATE_FILENAME = 'template.xlsx';
export const FINAL_FILENAME = 'final.xlsx';
export const VALUES_FILE_FULL_PATH = path.join(process.cwd(), 'files', VALUES_FILENAME);;
export const TEMPLATE_FULL_PATH = path.join(process.cwd(), 'files', TEMPLATE_FILENAME);
export const FINAL_FILE_FULL_PATH = path.join(process.cwd(), 'files', FINAL_FILENAME);

export const NUMBER_TO_ALPHABET = {
    0: 'A',
    1: 'B',
    2: 'C',
    3: 'D',
    4: 'E',
    5: 'F',
    6: 'G',
    7: 'H',
    8: 'I',
    9: 'J',
    10: 'K',
    11: 'L',
    12: 'M',
    13: 'N',
    14: 'O',
    15: 'P',
    16: 'Q',
    17: 'R',
    18: 'S',
    19: 'T',
    20: 'U',
    21: 'V',
    22: 'W',
    23: 'X',
    24: 'Y',
    25: 'Z',
};
