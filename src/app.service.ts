import { BadRequestException, Injectable } from '@nestjs/common';
import { Workbook, Worksheet } from 'exceljs'
import * as tmp from 'tmp-promise';
import { writeFile } from 'fs/promises'
import { Observable, catchError, from, lastValueFrom, switchMap, throwError } from 'rxjs';


@Injectable()
export class AppService {

  async getHello() {

    let rows = []

    let data = [
      {
        tsoli_id: 2389268,
        trepr_nombre: 'Consulado General de El Salvador en Doral, Florida, Estados Unidos',
        tsoli_fechahoracreacion: '2023-10-24T20:58:53.000Z',
        nombre: 'NEHEMIAS ABDIEL DE LEON MIRANDA',
        tsoli_concepto: null,
        cpais_nombre: 'United States of America',
        cserv_nombre: 'DUI por primera vez',
        cserv_precio: '35.00',
        tsoli_cantidaddocs: 1,
        catlin_nombre: 'EMISION DE DUI',
        treci_recibo: '10215',
        tpers_telefono: null
      },
      {
        tsoli_id: 2389269,
        trepr_nombre: 'Consulado General de El Salvador en Los Angeles, California, Estados Unidos',
        tsoli_fechahoracreacion: '2023-10-25T20:58:53.000Z',
        nombre: 'JOSE MANUEL RODRIGUEZ LOPEZ',
        tsoli_concepto: null,
        cpais_nombre: 'United States of America',
        cserv_nombre: 'DUI por primera vez',
        cserv_precio: '35.00',
        tsoli_cantidaddocs: 1,
        catlin_nombre: 'EMISION DE DUI',
        treci_recibo: '10216',
        tpers_telefono: null
      },
      {
        tsoli_id: 2389270,
        trepr_nombre: 'Consulado General de El Salvador en Houston, Texas, Estados Unidos',
        tsoli_fechahoracreacion: '2023-10-26T20:58:53.000Z',
        nombre: 'MARIA ISABEL GONZALEZ RAMIREZ',
        tsoli_concepto: null,
        cpais_nombre: 'United States of America',
        cserv_nombre: 'DUI por primera vez',
        cserv_precio: '35.00',
        tsoli_cantidaddocs: 1,
        catlin_nombre: 'EMISION DE DUI',
        treci_recibo: '10217',
        tpers_telefono: null
      }
    ];


    data.forEach((docs: any) => {
      rows.push(Object.values(docs));
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
    rows.unshift(Object.keys(data[0]));

    sheet.addRows(rows);

    // add style to the table
    this.styleSheet(sheet);

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
    try {
      // Create a temporary file
      const tmpFile = await tmp.file({
        discardDescriptor: true,
        prefix: 'myExcelSheetTest',
        postfix: '.xlsx',
        mode: parseInt('0600', 8)
      });

      // Write to the file using exceljs
      await book.xlsx.writeFile(tmpFile.path);

      console.log(tmpFile.path);

      // Return the file path
      return tmpFile.path;
    } catch (error) {
      throw new BadRequestException(error);
    }

    //function with rxjs
    // const file$ = from(tmp.file({
    //   discardDescriptor: true,
    //   prefix: 'myExcelSheetTest',
    //   postfix: '.xlsx',
    //   mode: 0o600
    // })).pipe(
    //   switchMap(tmpFile => from(book.xlsx.writeFile(tmpFile.path)).pipe(
    //     switchMap(() => from([tmpFile.path])),
    //     catchError(error => throwError(() => new Error(error)))
    //   )),
    //   catchError(error => throwError(() => new Error(error)))
    // )

    // return await lastValueFrom(file$)

  }

  private styleSheet(sheet: Worksheet) {

    //set the width of each column

    sheet.getColumn(1).width = 20.5
    sheet.getColumn(2).width = 20.5
    sheet.getColumn(3).width = 20.5
    sheet.getColumn(4).width = 20.5
    sheet.getColumn(5).width = 20.5
    sheet.getColumn(6).width = 20.5
    sheet.getColumn(7).width = 20.5
    sheet.getColumn(8).width = 20.5
    sheet.getColumn(9).width = 20.5
    sheet.getColumn(10).width = 20.5
    sheet.getColumn(11).width = 20.5
    sheet.getColumn(12).width = 20.5

    //set the height of header

    //font color
    sheet.getRow(1).height = 30.5
    sheet.getRow(2).height = 40.5
    // sheet.getRow(3).height = 30.5
    // sheet.getRow(4).height = 30.5
    // sheet.getRow(5).height = 30.5
    // sheet.getRow(6).height = 30.5
    // sheet.getRow(7).height = 30.5
    // sheet.getRow(8).height = 30.5
    // sheet.getRow(9).height = 30.5
    // sheet.getRow(10).height = 30.5
    // sheet.getRow(11).height = 30.5
    // sheet.getRow(12).height = 30.5

    //font color

    sheet.getRow(1).font = { size: 11.5, bold: true, color: { argb: 'FFFFFF' } }
    // sheet.getRow(2).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(3).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(4).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(5).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(6).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(7).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(8).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(9).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(10).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(11).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }
    // sheet.getRow(12).font = {size: 11.5, bold: true, color: {argb: 'FFFFFF'} }

    //background color

    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: '000000' }, fgColor: { argb: '000000' } }
    // sheet.getRow(2).fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(3).fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(4).fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(5).fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(6).fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(7).fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(8).fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(9).fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(10).fill  = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(11).fill  = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }
    // sheet.getRow(12).fill  = {type: 'pattern', pattern: 'solid', bgColor: {argb: '000000'}, fgColor: { argb: '000000' } }

    //alignments
    sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    sheet.getRow(2).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(3).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(4).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(5).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(6).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(7).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(8).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(9).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(10).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(11).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // sheet.getRow(12).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }

    //borders
    sheet.getRow(1).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(2).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(3).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(4).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(5).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(6).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(7).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(8).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(9).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(10).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(11).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }
    // sheet.getRow(12).border = { top: { style: 'thin', color: { argb: '000000' } }, left: { style: 'thin', color: { argb: 'FFFFFF' } }, bottom: { style: 'thin', color: { argb: '000000' } }, right: { style: 'thin', color: { argb: 'FFFFFF' } } }


  }


}
