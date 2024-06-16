import type { Sheet } from "@eyeseetea/xlsx-populate";

export type Either<A, B> = Readonly<[A, (null | undefined)?]> | Readonly<[(null | undefined), B]>;

export type Result<T, E extends Error = Error> = Either<E, T>;

export type PromiseResult<T, E extends Error = Error> = Promise<Result<T, E>>;

export type NewValue = {
    'Función': string;
    'Falla Funcional': string;
    'Modo de Falla (Causa de la Falla)': string[];
    'Efecto Inicial de la Falla (Que ocurre cuando Falla)': string[];
    'Efecto Final de la Falla o Consecuencia (Que ocurre cuando Falla)': string[];
    'Tipo de Consecuencia': string;
    'Tipo de Tareas': string;
    'Descripción de las Tareas Propuestas': string[];
    'Frecuencia Ej. [7D,2D,7M,2A,1A,3M]': string;
    'Ejecutor Ej. [Oper.,Mec.,Pred.,Inst.,Elec.]': string;
    'Cantidad de Ejecutantes': string;
    'Horas Hombre Ej. [1x1,2x4,4x2,2x2]': string;
    'Equipo Operando -> Si,No': string;
};

export type GroupedNewValues = Record<string, NewValue[]>;

export type CopyAndPasteRangeParams = {
    worksheet: Sheet;
    srcRowStart: number;
    srcColStart: number;
    srcRowEnd: number;
    srcColEnd: number;
    destRowStart: number;
    destColStart: number;
};
