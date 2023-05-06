1.Read list of commands from Excel Sheet.
2.Write all the above list of commands to serial port.
3.Read Command Response from EVK board(Qualcomm) via PySerial API.
4.Command Response is read using PySerial readline() api,and parsed to the new line to bifuracte next command response.
5.Each Command Response & the Expected output from the same input Excel sheet are compared.
6.Each test case result is validated and Pass/Fail is reported.

Examples
#i2cdbgr -D /dev/i2c1 -s 0x2A -r -b 1utf8

i2cdbgr:i2c_combined_writeread(/dev/i2c1= 3,slaveAddr = 0x2a,addr:0x0, data:0x0,failed)


#i2cdbgr /dev/i2c3 0x3c read 2 0x54 1
addr: 0x54 data : 0x80 0x00     

