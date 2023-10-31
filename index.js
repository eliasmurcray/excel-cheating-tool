const XLSX = require('xlsx');

const uploadButton = document.getElementById('upload');
const downloadButton = document.getElementById('download');
let workbook;

uploadButton.oninput = async () => {
  const file = uploadButton.files[0];
  if(!file) return;
  const data = await file.arrayBuffer();
  workbook = XLSX.read(data);

  let first_sheet_name = workbook.SheetNames[0];
  let worksheet = workbook.Sheets[first_sheet_name];

  const range = XLSX.utils.decode_range(worksheet['!ref']);
  for(let R = range.s.r; R <= range.e.r; ++R) {
    for(let C = range.s.c; C <= range.e.c; ++C) {
      let cellref = XLSX.utils.encode_cell({c: C, r: R});
      if(!worksheet[cellref]) continue;
      let cell = worksheet[cellref];
      let comment = cell.c[0].t.split('"')[3];
      if(comment.startsWith('=')) {
        cell.f = comment;
        cell.t = 'n';
      } else {
        cell.v = comment;
        if(isNumeric(comment)) {
          cell.t = 'n';
        } else {
          cell.t = 's';
        }
      }
      delete cell.c;

      console.log(cell);

      worksheet[cellref] = cell;
    }
  }

  console.log(worksheet);

  downloadButton.disabled = false;
};

function isNumeric(str) {
  if (typeof str != "string") return false // we only process strings!  
  return !isNaN(str) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
         !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
}

downloadButton.onclick = () => {
  XLSX.writeFile(workbook, 'output.xlsx');
};

/**{
    "a": "Keith Hekman",
    "t": "Formula: you had \"\" and the key had \"Using Matrix Equations to Solve\"\n",
    "r": "<r><rPr><b/><sz val=\"9\"/><color indexed=\"81\"/><rFont val=\"Tahoma\"/><charset val=\"1\"/></rPr><t xml:space=\"preserve\">Formula: you had \"\" and the key had \"Using Matrix Equations to Solve\"\r\n</t></r>",
    "T": false,
    "h": "<span style=\"font-size:9pt;\"><b>Formula: you had \"\" and the key had \"Using Matrix Equations to Solve\"<br/></b></span>"
} */