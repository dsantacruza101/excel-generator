import { Controller, Get, Header, Res } from '@nestjs/common';
import { AppService } from './app.service';
import { Response } from 'express';

@Controller('excel')
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get()
  @Header('Content-Type', 'text/xlsx')
  async getHello(
    @Res() res: Response
  ):Promise<any> {
    const data = await this.appService.getHello(); 
    console.log(data);
    
    res.download(`${data}`)
  }
}
