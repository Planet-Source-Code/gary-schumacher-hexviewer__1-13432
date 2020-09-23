// HexView.cpp 
//Used to increase speed of loop in Visual Basic HexViewer Program
//This section will benefit the most from a C++ optimization
//The Visual Basic statements are included to provide a tutorial
//for the C++ challenged among us (mostly me).
//
//Opted not to use strncpy for moving small #'s of bytes..more overhead			
		//strncpy( tabChar,HexTable+cOneChar*2,2);

//I credit an article by Jonathan Morrison in Visual Basic Programmer's Journal Jan/2001
//for the DLL technique used.  "Add C++ to your VB Apps" 


//************ Tip for debugging Dll Routines that are called from Visual Basic **********
//This applies to Visual Studio 6.0

//1.  In Project| Settings...  Select Debug Tab
//2.  Under "Executable for Debug Session:" Browse and enter your VB6.Exe file
//3.  Set Breakpoint in Dll
//4.  In Build| Start Debug...Select Go
//5.  This will start Visual Basic...run your Dll Calling program...

//One note...when I was trying to use small __asm blocks, I was getting some
//strange errors.  Don't know if this was a bug or what.    
			

#include "stdafx.h"
#include "HexView.h"
#include "ctype.h"		//isAlpha and isDigit Functions
#include "string.h"		//String Functions
#include "stdio.h"		//sprintf function

//The Following offsets are used for the formatted output record
#define LeftHexOff  10              //Offset of 1st Hex char in strWork (Left side)
#define RightHexOff  36             //Offset of 8th Hex char in strWork (Right side
#define AsciiOff  62                //Offset of 1st Ascii value in strWork
#define LenDummy  79				//Length of Dummy Formatted String Record

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}

void APIENTRY HexDump(char *source, char *dest, long filelen, long lNumRecs)
{
	unsigned char cOneChar;				//Byte(Char) to be converted
	unsigned char cHexValue[2];			//Byte to 2 Digit Hex conversion... 5D = "5D"
	long lOutputOff=0;					//Keep track of position in Output String 
	

	//Fill entire Output String with Dummy Format Records
	for (long lFillDest=0; lFillDest < lNumRecs*LenDummy; lFillDest=lFillDest+LenDummy)
	{	
		memcpy(dest+lFillDest,strDummyOutput,79);
	}
	//Create formatted output from input file
	for (long lSourceOff=0; lSourceOff < filelen; lSourceOff=lSourceOff+16)
	{
	   	
		//Translate Offset (lSourceOff) to an 8 byte Hex Address
		sprintf(dest+lOutputOff,"%08X",lSourceOff);
		//Overlay 0x00 from sprintf translation with colon
		dest[lOutputOff+8]=':';
		
		//For intPtrSrce = 0 To 15                'Point to each location in source record (m_strSourceFile)
		//Loop through each character in the string
		for (int i=0; i < 16; i++)
		{
			//     bytAscii = Asc(Mid(m_strSourceFile, lngSourceOffset + intPtrSrce, 1)) 'Extract char from source file
			cOneChar=source[lSourceOff+i];			   //Extract char to be xlated to hex
			
			//     strTemp = m_arrXlate(bytAscii, 1)   'Convert to 2 byte hex value from table
			cHexValue[0]=HexTable[cOneChar*2];
			cHexValue[1]=HexTable[cOneChar*2+1];
			//     If intPtrSrce <= 7 Then             'Working on Left side of Output
			if (i <= 7)
			{
				//         Mid(strWork, LeftHexOff + intPtrSrce * 3, 2) = strTemp
				dest[lOutputOff+LeftHexOff+i*3]=cHexValue[0];
				dest[lOutputOff+LeftHexOff+(i*3)+1]=cHexValue[1];	
			}
			else
			{
				//         Mid(strWork, RightHexOff + (intPtrSrce - 8) * 3, 2) = strTemp
				dest[lOutputOff+RightHexOff+((i-8)*3)]=cHexValue[0];
				dest[lOutputOff+RightHexOff+((i-8)*3)+1]=cHexValue[1];	
			}

			//Is the character an Upper/Lower Case letter or Number?
			if (isalpha(cOneChar) || isdigit(cOneChar))
			{
      			//      Mid(strWork, AsciiOff + intPtrSrce, 1) = Chr$(bytAscii)
				dest[lOutputOff+AsciiOff+i]=cOneChar;
				
			}
		}					//for (int i=0; i < 15; i++) loop
		lOutputOff=lOutputOff+LenDummy;		//Next output record to format
	}						//Main Loop
}							//HexDump Apientry







