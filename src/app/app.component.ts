import { Component } from '@angular/core';
import ExcelJS, { Workbook } from 'exceljs';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'TestExcel';

  // carico il file
  async uploadFile($event: any) {

    const fileInput = $event.target.files[0];
    const fileReader = new FileReader();

    fileReader.readAsArrayBuffer(fileInput);

    fileReader.onload = async () => {
      this.writeExcel(fileReader.result);
    };
  }

  // genero l'oggetto workbook
  async writeExcel(buffer) {
    const wb = new ExcelJS.Workbook();

    wb.xlsx.load(buffer).then(workbook => {
      this.modifyExcel(workbook);
    });
  }

  // effettuo le modifiche
  async modifyExcel(workbook: Workbook) {

    const worksheet = workbook.getWorksheet('pianificazione');

    // modifiche
    let rowStartIndex = 3;
    let colStartIndex = 65; // A

    // dati di test
    let testData = [
      {id: 1, tracking: 'T', norad: '29165', date: '02 August 2021', start: '18:57:52', stop: '19:17:07', durata: '00:19:15', delivery: '', note: '' },
      {id: 2, tracking: 'S', norad: '43637', date: '03 August 2021', start: '19:22:18', stop: '19:26:03', durata: '00:03:45', delivery: '', note: '' },
      {id: 3, tracking: 'T', norad: '12446', date: '04 August 2021', start: '19:33:09', stop: '19:53:24', durata: '00:20:15', delivery: '', note: '' },
    ]

    let rowIndex = rowStartIndex;
    let colIndex = colStartIndex;
    let cellPos = '';

    const highlightStyle: any = { type: 'pattern', pattern:'solid', fgColor: {argb:'FFFF00'} };
    const borderStyle: any = { top: {style:'thin', color: {argb:'000000'}}, left: {style:'thin', color: {argb:'000000'}}, bottom: {style:'thin', color: {argb:'000000'}}, right: {style:'thin', color: {argb:'000000'}} };
    const errorStyle: any = { color: { argb: 'ff0000' } };
    const alignmentStyle: any = { vertical: 'middle', horizontal: 'center' };

    testData.forEach(element => {
      colIndex = colStartIndex;

      // scrivi: id
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.id.toString();
      worksheet.getCell(cellPos).fill = highlightStyle;
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      colIndex++; // avanzamento di colonna

      // scrivi: tracking
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.tracking;
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      colIndex++; // avanzamento di colonna

      // scrivi: norad
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.norad.toString();
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      worksheet.getCell(cellPos).font = errorStyle;
      colIndex++; // avanzamento di colonna

      // scrivi: date
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.date;
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      colIndex++; // avanzamento di colonna

      // scrivi: start
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.start;
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      colIndex++; // avanzamento di colonna

      // scrivi: stop
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.stop;
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      colIndex++; // avanzamento di colonna

      // scrivi: durata
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.durata;
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      colIndex++; // avanzamento di colonna

      // scrivi: delivery
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.delivery;
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      colIndex++; // avanzamento di colonna

      // scrivi: note
      cellPos = String.fromCharCode(colIndex) + rowIndex.toString();
      worksheet.getCell(cellPos).value = element.note;
      worksheet.getCell(cellPos).border = borderStyle;
      worksheet.getCell(cellPos).alignment = alignmentStyle;
      colIndex++; // avanzamento di colonna

      // avanzamento di riga
      rowIndex++;
    });

    // nome del sensore
    worksheet.getCell('A1').value = 'BIRALES';

    // export
    this.exportExcel(workbook);
  }

  // esporto il file
  async exportExcel(workbook) {
    let bufferModified = await workbook.xlsx.writeBuffer();

    const fileName = "NOMESENSORE_20210802T1857_20210803T0315_PIANIFICAZIONE.xlsx";
    const fileBlob = new Blob([bufferModified], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    this.fromBlobToFile(fileBlob, fileName);
  }

  fromBlobToFile(blob:Blob, fileName:string){
    var a = document.createElement("a");
    document.body.appendChild(a);

    var url = window.URL.createObjectURL(blob);
    a.href = url;
    a.download = fileName;
    a.click();
    window.URL.revokeObjectURL(url);
    a.remove();
  }
}
