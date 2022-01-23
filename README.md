# SMA.NET
This is a small Visual Basic .NET project for communicating with SMA Sunnyboy solar panel inverters. It uses the RS485 interface at 1200 baud, it is designed to run on the Raspberry Pi but will run in Windows too.

# Dependencies
The project is written entirely in .NET and doesn't depend on other libraries like Yasdi. The only requirement is to install mono on the Raspberry pi to be able to run the SMAReader.exe file.

# Installation
In Linux/Raspberry Pi, first install mono

```
sudo apt update
sudo apt install mono-runtime mono-vbnc
```

Download and compile the source with Visual Studio or download the realease zip file and extract the SMADOTNET.dll and SMAReader.exe files.
On Windows you can just run the SMAReader.exe from the command line, in Linux/Raspberry Pi start the executable with mono.

```
sudo mono SMAReader.exe
```

Raspberry Pi users need to connect the TX/RX pin to a converter IC like the MAX487. Since RS485 is half-duplex you'll also need a pin to switch the MAX487 between reading or writing. By default gpio 4 is used for this. Alternative is to use a USB to RS485 converter.

# Reading data
You'll need to know the network address of your Sunnyboy inverter. If it's the first time you've ever connected you'll probably need to configure the address. If you have already connected other software with the converter the network address is probably already configured.

## I know the network address
If you know the network address you can direct get the current values with the **get** command. Add -h parameter to get more human readable output, the default is json.

```
sudo mono SMAReader.exe get -a 201
```

By default it will use the Raspberry Pi build in COM port, to use a different port (like USB to RS485) add the -d parameter. For example:

```
sudo mono SMAReader.exe get -d /dev/ttyUSB0 -a 201
```

## I don't know network address
If you don't know the network address but you know the serialnumber you can use the **find** command to search for it.

```
sudo mono SMAReader.exe find -s 2000506234
```

If you don't get a reply the network address probably isn't configured yet. Try the **unconfigured** command to list all Sunnyboy devices that do not have a network address. It will return the serialnumbers.

```
sudo mono SMAReader.exe unconfigured
```

Next you can assign a network address for the serial number found with the **config** command.

```
sudo mono SMAReader.exe -s 2000506234 -a 201
```
Where 201 is the network address you want to assign, pick a unique number between 1 and 65535 as address for each Sunnyboy connected to the RS485 bus.

# Build your own application
With the SMADOTNET.dll file you can also build your own applications in C# or VB.NET. For example to save to SQL or upload with HTTP.

# Remote debugging
You can use Visual Studio Code and the mono debug extension to debug the application. For this run the mono application with extra parameters:

```
sudo mono --debug --debugger-agent=transport=dt_socket,server=y,address=192.168.0.227:55555 SMAReader.exe get -a 201
```

In the SMA.NET project directory create a folder **.vscode**, inside this folder create a file called **launch.json**. Open this **launch.json** file and paste text:

```
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Attach to Mono",
            "request": "attach",
            "type": "mono",
            "address": "192.168.0.227",
            "port": 55555
        }
    ]
}
```

Where 192.168.0.227 is the IP of your Raspberry. Now open the folder with Visual Studio code (righ click open with Code). Go to the run->start debugging (F5) menu and choose "Attach to Mono". Now the mono will start running on your Raspberry Pi and you can set breakpoints etc...

NOTE: 
- If you recompile the source code don't forget to set **Generate debug info** parameter to **portable** in the **Advanced compile options** menu of Visual Studio. 
- Also copy the .pdb files to the Raspberry Pi if you want to remote debug.