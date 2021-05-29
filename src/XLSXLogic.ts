import XLSX from 'xlsx'

export enum ErrorCode
{
    NONE,
    MISSING_INPUT,
    MISSING_TEMPLATE,
    UNKNOWN
}

export function validateInput(inputFile : FileList, templateFile : FileList) : ErrorCode
{
    if(inputFile === null)
    {
        return ErrorCode.MISSING_INPUT;
    }
    if(templateFile === null)
    {
    	return ErrorCode.MISSING_TEMPLATE
    }
    return ErrorCode.NONE;
}

export function matchFiles(inputFile : File, templateFile : File) : Promise<void>
{
    return new Promise((resolve, _reject)=>{
        Promise.all([
            readXLSXFile(inputFile),
            readXLSXFile(templateFile)]
        ).then(values=>{
            resolve(matchXLSXs(values[0], values[1]));
        });
    });
}

function matchXLSXs(input : XLSX.WorkBook, template : XLSX.WorkBook)
{
    XLSX.writeFile(input, "output.xlsx");
}

function readXLSXFile(file : File) : Promise<XLSX.WorkBook>
{
    return new Promise((resolve, _reject) => {
        var reader = new FileReader();
        reader.onload = function(e) {
            var workbook = XLSX.read(e.target.result, {type: 'array'});
            resolve(workbook);
        };
        reader.readAsArrayBuffer(file);
    })
}

// function parseXLSX(workbook : XLSX.WorkBook)
// {
//     // for(var R = range.s.r; R <= range.e.r; ++R) {
//     // 	for(var C = range.s.c; C <= range.e.c; ++C) {
//     // 		var cell_address = {c:C, r:R};
//     // 		/* if an A1-style address is needed, encode the address */
//     // 		var cell_ref = XLSX.utils.encode_cell(cell_address);
//     // 	}
//     // }
// }