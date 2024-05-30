import { BadRequestException, Injectable } from '@nestjs/common';
import { Workbook, Worksheet } from 'exceljs'
import * as tmp from 'tmp-promise';
import { writeFile } from 'fs/promises'
import { Observable, catchError, from, lastValueFrom, switchMap, throwError } from 'rxjs';
import { HeaderTittles } from './common/consolidado-af-goes.structure';
import { data } from './database/data';

@Injectable()
export class AppService {

  async getHello() {

    let rows = []

    data.forEach((_data: any, index: number) => {
      rows.push([index + 1,
      _data.codigo,
      _data.nombre,
      _data.cantidad,
      _data.valor,
      _data.cargos.cantidad,
      _data.cargos.valor,
      _data.descargos.cantidad,
      _data.descargos.valor,
      _data.decrimento.cantidad,
      _data.decrimento.valor,
      _data.suma_por_ajuste.cantidad,
      _data.suma_por_ajuste.valor,
      _data.resta_por_ajuste.cantidad,
      _data.resta_por_ajuste.valor]);
    });

    //creating a workbook
    let book = new Workbook();

    book.creator = 'daniel'

    // creating a worksheet to workbook
    let sheet = book.addWorksheet('sheet', {
      pageSetup: { fitToPage: true, fitToHeight: 5, fitToWidth: 7, paperSize: undefined },
      headerFooter: { firstHeader: 'Hello World ExcelJS', firstFooter: 'Good Bye World ExcelJS' },

    });

    // add the header
    // rows.unshift(Object.keys(data[0]));
    console.log(rows);

    // add style to the table
    this.style_AF_Sheet(sheet, rows);

    // const matrixSheet = book.addWorksheet('Matrix');

    // const matrix = [
    //   [1, 2, 3],
    //   [4, 5, 6],
    //   [7, 8, 9]
    // ];

    // for ( let i = 0; i < matrix.length; i++ ) {
    //   for(let j = 0; j < matrix[i].length; j++){
    //     matrixSheet.getCell(i + 1, j + 1).value = matrix[i][j];
    //   }
    // }

    // Original function:
    // let filePromise = await new Promise((resolve, reject) => {
    //   tmp.file({
    //     discardDescriptor: true,
    //     prefix: 'myExcelSheetTest',
    //     postfix: '.xlsx',
    //     mode: parseInt('0600', 8)
    //   }, async (err, file) => {
    //     if (err) {
    //       throw new BadRequestException(err);
    //     }

    //     book.xlsx.writeFile(file).then(() => {
    //       resolve(file)
    //       console.log(file);

    //     }).catch(error => {
    //       throw new BadRequestException(error);
    //     })
    //   })
    // });

    // return filePromise;

    // updated function-must install tmp-promise
    // try {
    //   // Create a temporary file
    //   const tmpFile = await tmp.file({
    //     discardDescriptor: true,
    //     prefix: 'myExcelSheetTest',
    //     postfix: '.xlsx',
    //     mode: parseInt('0600', 8)
    //   });

    //   // Write to the file using exceljs
    //   await book.xlsx.writeFile(tmpFile.path);

    //   console.log(tmpFile.path);

    //   // Return the file path
    //   return tmpFile.path;
    // } catch (error) {
    //   throw new BadRequestException(error);
    // }

    //function with rxjs
    const file$ = from(tmp.file({
      discardDescriptor: true,
      prefix: 'myExcelSheetTest',
      postfix: '.xlsx',
      mode: 0o600
    })).pipe(
      switchMap(tmpFile => from(book.xlsx.writeFile(tmpFile.path)).pipe(
        switchMap(() => from([tmpFile.path])),
        catchError(error => throwError(() => new Error(error)))
      )),
      catchError(error => throwError(() => new Error(error)))
    )

    return await lastValueFrom(file$)

  }

  public style_AF_Sheet(sheet: Worksheet, dataRows?: any) {

    const tittleCells = ['A1:Q1', 'A2:Q2', 'A4:Q4', 'A6:E6'];

    tittleCells.forEach(range => {
      sheet.mergeCells(range)
    });

    const tittleValues = HeaderTittles.TITTLE_CONSOLIDADO_AF_GOES;

    tittleValues.forEach(({ cell, value, fontSize }) => {
      this.setTittleCells(sheet, cell, value, fontSize)
    });

    const consAfGoes = HeaderTittles.HEADER_ROW_CONSOLIDADO_AF_GOES;

    consAfGoes.forEach(({ range, value }) => {
      this.formatCellsForheaderTable(sheet, range, value)
    })

    
    this.setBorders(sheet, 7, 1, 9, 15);
    
    sheet.addRows(dataRows);

  }

  public setTittleCells = (sheet: Worksheet, cell: string, value: string, fontSize: number) => {
    const titleCell = sheet.getCell(cell);
    titleCell.value = value;
    titleCell.style.font = { bold: true, size: fontSize };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  }

  public formatCellsForheaderTable = (sheet: Worksheet, cellRange: string, value: string) => {

    sheet.mergeCells(cellRange);
    const cell = sheet.getCell(cellRange.split(':')[0]);
    cell.value = value;
    cell.style.font = { size: 9 };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

  }

  public setBorders = (sheet, startRow: number, startCol: number, endRow: number, endCol: number) => {
    const borderStyle = { style: 'thin' };

    for (let i = startRow; i <= endRow; i++) {
      for (let j = startCol; j <= endCol; j++) {
        const cell = sheet.getCell(i, j);
        cell.border = {
          top: borderStyle,
          left: borderStyle,
          bottom: borderStyle,
          right: borderStyle
        };
      }
    }
  }

}
