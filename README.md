# CST Microwave Studio Interaction
CST STUDIO SUITE is implemented as OLE automation server. Firstly it is nessesary to set connection. For example:
```sh
Type cstAppType = Type.GetTypeFromProgID("CSTStudio.Application.2015");
object cstApp = Activator.CreateInstance(cstAppType);
object cstDocument = cstAppType.InvokeMember("NewMWS", BindingFlags.InvokeMethod, null, cstApp, new object[] { });
```
Units settings example with comparison between MWS history and C# code

**MWS History list item:**
```sh
With Units 
     .Geometry "mm" 
     .Frequency "GHz" 
     .Time "ns" 
     .TemperatureUnit "Kelvin" 
     .Voltage "V" 
     .Current "A" 
     .Resistance "Ohm" 
     .Conductance "Siemens" 
     .Capacitance "PikoF" 
     .Inductance "NanoH" 
End With 
```
**C#**
```sh
object units = cstAppType.InvokeMember("Units", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
cstAppType.InvokeMember("Geometry", BindingFlags.InvokeMethod, null, units, new object[] { "mm" });
cstAppType.InvokeMember("Frequency", BindingFlags.InvokeMethod, null, units, new object[] { "GHz" });
cstAppType.InvokeMember("Time", BindingFlags.InvokeMethod, null, units, new object[] { "ns" });
cstAppType.InvokeMember("TemperatureUnit", BindingFlags.InvokeMethod, null, units, new object[] { "Kelvin" });
cstAppType.InvokeMember("Voltage", BindingFlags.InvokeMethod, null, units, new object[] { "V" });
cstAppType.InvokeMember("Current", BindingFlags.InvokeMethod, null, units, new object[] { "A" });
cstAppType.InvokeMember("Resistance", BindingFlags.InvokeMethod, null, units, new object[] { "Ohm" });
cstAppType.InvokeMember("Conductance", BindingFlags.InvokeMethod, null, units, new object[] { "Siemens" });
cstAppType.InvokeMember("Capacitance", BindingFlags.InvokeMethod, null, units, new object[] { "PikoF" });
cstAppType.InvokeMember("Inductance", BindingFlags.InvokeMethod, null, units, new object[] { "NanoH" });
```

To find description for all available objects use VBA manual included.
