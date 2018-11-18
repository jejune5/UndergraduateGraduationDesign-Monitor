#include "stc12c5a60s2.h"
#include "intrins.h"
#define uchar unsigned char
#define uint unsigned int
#define u8 unsigned char
#define u16 unsigned int
#define vu8 unsigned char
#define vu16 unsigned int
#define vu32 unsigned long int
sbit Key1=P3^7;
sbit Key2=P3^6;
sbit Key3=P3^5;
void delayms(int z)
{
	int x,y;
	for(x=110;x>0;x--)
		for(y=z;y>0;y--);
}

vu8 sys[20];
#define temph sys[1]
#define templ sys[2]
#define adch sys[3]
#define adcl sys[4]
#define shuiweih sys[5]
#define shuiweil sys[6]
#define bilih  sys[7]
#define bilil1  sys[8]
#define highh_h sys[9]
#define highh_l sys[10]


#include "1602_12864.h"
#include "usart.h"	   //串行通讯接口
#include "dht11.h"
vu8 UartRx[35];
vu16 RecCnt;
vu16 RecTimeCount;
vu8 UartRecInit=0,Uart_RecOk=0,UartFlag_RecTime=0;

void Timer0_Init(void)
{
	//定时器0初始化
	TH0=(65536-10000)/256;
	TH0=(65536-10000)%256;
	TMOD|=0x01;
	EA=1;
	ET0=1;
	TR0=1;
}
					  
sfr ADC_LOW2    =   0xBE;           //ADC low 2-bit result register
/*Define ADC operation const for ADC_CONTR*/
#define ADC_POWER   0x80            //ADC power control bit
#define ADC_FLAG    0x10            //ADC complete flag
#define ADC_START   0x08            //ADC start control bit
#define ADC_SPEEDLL 0x00            //420 clocks
#define ADC_SPEEDL  0x20            //280 clocks
#define ADC_SPEEDH  0x40            //140 clocks
#define ADC_SPEEDHH 0x60            //70 clocks

void InitADC(void)
{
    P1ASF |= ((1<<2));   //P12,P13作为ADC口 
    ADC_RES = 0;                    //Clear previous result
    ADC_CONTR = ADC_POWER | ADC_SPEEDLL;
    Delay_NMS(2);                       //ADC power-on and delay
}

//获取ADC的值
uint GetADCResult(uchar ch)
{
	uint i;
    ADC_CONTR = ADC_POWER | ADC_SPEEDLL | ch | ADC_START;
    _nop_();                        //Must wait before inquiry
    _nop_();
    _nop_();
    _nop_();
    while (!(ADC_CONTR & ADC_FLAG));//Wait complete flag
    ADC_CONTR &= ~ADC_FLAG;         //Close ADC
	i=ADC_RES;
	i=(i<<2)|ADC_LOW2;
    return i;                 //Return ADC result
}

void main(void)//切换界面才进行保存
{

    vu8 i,j;
    LCD_Init();
    SM0=0;
    SM1=1;
    EA=1;
    TMOD=0x21;
    REN=1;

    TH1=0xfd;
    TL1=0xfd;
    TR1=1;
    
    TH0=100;
    TL0=100;
    EA=1;
    TR0=1;
    ET0=1;
   
    while(1)
    {
        DH11_GetTempDamp();

        LCD_DisStr(0,0,"Temp:");
        LCDW_Dat(TempNow/10+48);
        LCDW_Dat(TempNow%10+48);
        LCD_DisStr(0,9,"Damp:");
        LCDW_Dat(DampNow/10+48);
        LCDW_Dat(DampNow%10+48);
        i=GetADCResult(2)/1024.0*5.0*10.0;
      
        LCD_DisStr(1,0,"Light:");
        LCDW_Dat(i/10+48);
        LCDW_Dat(i%10+48);       
        SBUF='S';
        while(!TI);
        TI=0;   
      
SBUF=TempNow;
        while(!TI);
        TI=0;
        
        SBUF=TempNow;
        while(!TI);
        TI=0;
       
         SBUF=DampNow;
        while(!TI);
        TI=0;
              SBUF=i;
        while(!TI);
        TI=0;       
              SBUF='E';
        while(!TI);
        TI=0;          
      delayms(888);
    }
}

/*************UART_1 中断服务程序****************/
void Uart1() interrupt 4 using 1
{
    if (RI)
    {
        RI = 0;                 //清除RI位
		if(UartRecInit==0)
		{
			UartRecInit=1;
			RecCnt=0;//接收个数
			UartFlag_RecTime=1;//允许接收倒计时
			RecTimeCount=10;//接收倒计时时间,1秒
			
		}
		if(RecCnt<32)	
			UartRx[RecCnt++]=SBUF;

		if(RecCnt>=32)

		{
			UartRecInit=0;
			RecTimeCount=100;//防程序偶合,没多大意义
			UartFlag_RecTime=0;//禁止接收倒计时	
			Uart_RecOk=1;
            ES=0;
		}       
    }
}

void Timer0() interrupt 1//100us中断
{
 	TH0=(65536-3000)/256;
 	TL0=(65536-3000)%256;
    if(UartFlag_RecTime)//从接收第一个有效字节开始,进行一秒的倒计时,如果没收到完整的
                //复位接收器,并发送一个字节的报错数据0x88.
    {
        if(RecTimeCount>0)
            RecTimeCount--;
        else
        {
            UartRecInit=0;
            UartFlag_RecTime=0;//禁止接收倒计时	
            Uart_RecOk=1;
            ES=0;
        }
    }
}


