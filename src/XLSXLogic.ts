import XLSX from 'xlsx'

export enum ErrorCode {
    NONE,
    MISSING_INPUT,
    MISSING_TEMPLATE,
    UNKNOWN
}

export function validateInput(inputFile: FileList, templateFile: FileList): ErrorCode {
    if (inputFile === null) {
        return ErrorCode.MISSING_INPUT;
    }
    if (templateFile === null) {
        return ErrorCode.MISSING_TEMPLATE
    }
    return ErrorCode.NONE;
}

export function matchFiles(inputFile: File, templateFile: File): Promise<void> {
    return new Promise((resolve, _reject) => {
        Promise.all([
            readXLSXFile(inputFile),
            readXLSXFile(templateFile)]
        ).then(values => {
            resolve(matchXLSXs(values[0], values[1]));
        });
    });
}

function matchXLSXs(input: XLSX.WorkBook, template: XLSX.WorkBook) {
    let templateSheet = template.Sheets[template.SheetNames[0]];
    let inputSheet = input.Sheets[input.SheetNames[0]];

    let templateRange: XLSX.Range = XLSX.utils.decode_range(templateSheet['!ref']);

    var maxRow = 1;
    var sourceCellAddress: XLSX.CellAddress = null;
    var destCellAddress: XLSX.CellAddress = null;
    for (var c = 0; c <= templateRange.e.c; c++) {
        destCellAddress = { c: c, r: 0 };
        console.log("matchXLSXs - iterate column: " + JSON.stringify(templateSheet[XLSX.utils.encode_cell(destCellAddress)]));
        sourceCellAddress = findColumnID(templateSheet[XLSX.utils.encode_cell(destCellAddress)], inputSheet);
        if (sourceCellAddress)
            maxRow = Math.max(maxRow, copyColumn(sourceCellAddress.c, destCellAddress.c, inputSheet, templateSheet));
    }

    var range = XLSX.utils.decode_range(templateSheet['!ref']);
    range.e.r = maxRow;
    templateSheet['!ref'] = XLSX.utils.encode_range(range);
    XLSX.writeFile(template, "output.xlsx");
}

function findColumnID(sourceCell: XLSX.CellObject, input: XLSX.WorkSheet): XLSX.CellAddress {
    if (!sourceCell)
        return null;

    let inputRange: XLSX.Range = XLSX.utils.decode_range(input['!ref']);
    var destCellAddress: string;
    for (var c = 0; c <= inputRange.e.c; c++) {
        destCellAddress = XLSX.utils.encode_cell({ c: c, r: 0 });
        console.log("findColumnID: " + JSON.stringify(input[destCellAddress]) + "  " + JSON.stringify(sourceCell));
        if (input[destCellAddress] && input[destCellAddress].w == sourceCell.w) {
            console.log("findColumnID: match");
            return { c: c, r: 0 };
        }
    }
    return null;
}

function copyColumn(sourceColumn: number, destColumn: number, input: XLSX.WorkSheet, template: XLSX.WorkSheet): number {
    var maxRow = 1;
    let inputRange: XLSX.Range = XLSX.utils.decode_range(input['!ref']);

    var dataCell: XLSX.CellObject = null;
    for (var r = 1; r <= inputRange.e.r; r++) {
        dataCell = input[XLSX.utils.encode_cell({ c: sourceColumn, r: r })]
        if (dataCell) {
            template[XLSX.utils.encode_cell({ c: destColumn, r: r })] = dataCell;
            maxRow = r;
        }
    }

    return maxRow;
}

function readXLSXFile(file: File): Promise<XLSX.WorkBook> {
    return new Promise((resolve, _reject) => {
        var reader = new FileReader();
        reader.onload = function (e) {
            var workbook = XLSX.read(e.target.result, { type: 'array' });
            resolve(workbook);
        };
        reader.readAsArrayBuffer(file);
    })
}