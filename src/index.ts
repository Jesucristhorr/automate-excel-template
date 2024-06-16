import { TEMPLATE_FULL_PATH, FINAL_FILE_FULL_PATH } from './config/constants';
import { Lib } from './lib';

import XlsxPopulate from '@eyeseetea/xlsx-populate';

(async () => {
    try {
        const [getValuesError, values] = await Lib.Value.getValuesToFill();

        if (getValuesError) {
            console.error('Get Values Error:', getValuesError);
            throw getValuesError;
        }

        const workbook = await XlsxPopulate.fromFileAsync(TEMPLATE_FULL_PATH);

        const worksheet = workbook.sheet('Sistema de Lubricación');

        if (!worksheet) throw new Error('No worksheet.');

        Lib.Template.fillData(values, worksheet);

        await workbook.toFileAsync(FINAL_FILE_FULL_PATH);
    } catch (error) {
        console.error('FATAL ERROR:', error);
        process.exit(1);
    }

    process.exit(0);
})();
