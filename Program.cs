using System;
using System.Reflection;

namespace Microwave_Studio_Interaction
{
    class Program
    {
        static void Main(string[] args)
        {
            //Connection
            Type cstAppType = Type.GetTypeFromProgID("CSTStudio.Application.2015");
            object cstApp = Activator.CreateInstance(cstAppType);
            object cstDocument = cstAppType.InvokeMember("NewMWS", BindingFlags.InvokeMethod, null, cstApp, new object[] { });

            //Here we are to change settings of project(background, units, etc)
            object background = cstAppType.InvokeMember("Background", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("ResetBackground", BindingFlags.InvokeMethod, null, background, new object[] { });
            cstAppType.InvokeMember("XminSpace", BindingFlags.InvokeMethod, null, background, new object[] { "0.0" });
            cstAppType.InvokeMember("XmaxSpace", BindingFlags.InvokeMethod, null, background, new object[] { "0.0" });
            cstAppType.InvokeMember("YminSpace", BindingFlags.InvokeMethod, null, background, new object[] { "0.0" });
            cstAppType.InvokeMember("YmaxSpace", BindingFlags.InvokeMethod, null, background, new object[] { "0.0" });
            cstAppType.InvokeMember("ZminSpace", BindingFlags.InvokeMethod, null, background, new object[] { "0.0" });
            cstAppType.InvokeMember("ZmaxSpace", BindingFlags.InvokeMethod, null, background, new object[] { "0.0" });
            cstAppType.InvokeMember("ApplyInAllDirections", BindingFlags.InvokeMethod, null, background, new object[] { "False" });


            object Material = cstAppType.InvokeMember("Background", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("Reset", BindingFlags.InvokeMethod, null, background, new object[] { });
            cstAppType.InvokeMember("Type", BindingFlags.InvokeMethod, null, background, new object[] { "Normal" });
            cstAppType.InvokeMember("Epsilon", BindingFlags.InvokeMethod, null, background, new object[] { "1.0" });
            cstAppType.InvokeMember("Mue", BindingFlags.InvokeMethod, null, background, new object[] { "1.0" });
            cstAppType.InvokeMember("ThermalType", BindingFlags.InvokeMethod, null, background, new object[] { "Normal" });


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


            //Let's draw something
            object materialCopper = cstAppType.InvokeMember("Material", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("Reset", BindingFlags.InvokeMethod, null, materialCopper, new object[] { });
            cstAppType.InvokeMember("Name", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "Copper (annealed)" });
            cstAppType.InvokeMember("Folder", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "" });
            cstAppType.InvokeMember("FrqType", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "static" });
            cstAppType.InvokeMember("Type", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "Normal" });
            cstAppType.InvokeMember("SetMaterialUnit", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "Hz", "mm" });
            cstAppType.InvokeMember("Epsilon", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "1" });
            cstAppType.InvokeMember("Mue", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "1.0" });
            cstAppType.InvokeMember("Kappa", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "5.8e+007" });
            cstAppType.InvokeMember("TanD", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0.0" });
            cstAppType.InvokeMember("TanDFreq", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0.0" });
            cstAppType.InvokeMember("TanDGiven", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "False" });
            cstAppType.InvokeMember("TanDModel", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "ConstTanD" });
            cstAppType.InvokeMember("KappaM", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0" });
            cstAppType.InvokeMember("TanDM", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0.0" });
            cstAppType.InvokeMember("TanDMFreq", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0.0" });
            cstAppType.InvokeMember("TanDMGiven", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "False" });
            cstAppType.InvokeMember("TanDMModel", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "ConstTanD" });
            cstAppType.InvokeMember("DispModelEps", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "None" });
            cstAppType.InvokeMember("DispModelMue", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "None" });
            cstAppType.InvokeMember("DispersiveFittingSchemeEps", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "Nth Order" });
            cstAppType.InvokeMember("DispersiveFittingSchemeMue", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "Nth Order" });
            cstAppType.InvokeMember("UseGeneralDispersionEps", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "False" });
            cstAppType.InvokeMember("UseGeneralDispersionMue", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "False" });
            cstAppType.InvokeMember("FrqType", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "all" });
            cstAppType.InvokeMember("Type", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "Lossy metal" });
            cstAppType.InvokeMember("SetMaterialUnit", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "GHz", "mm" });
            cstAppType.InvokeMember("Mue", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "1.0" });
            cstAppType.InvokeMember("Kappa", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "5.8e+007" });
            cstAppType.InvokeMember("Rho", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "8930.0" });
            cstAppType.InvokeMember("ThermalType", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "Normal" });
            cstAppType.InvokeMember("ThermalConductivity", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "401.0" });
            cstAppType.InvokeMember("HeatCapacity", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0.39" });
            cstAppType.InvokeMember("MetabolicRate", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0" });
            cstAppType.InvokeMember("BloodFlow", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0" });
            cstAppType.InvokeMember("VoxelConvection", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0" });
            cstAppType.InvokeMember("MechanicsType", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "Isotropic" });
            cstAppType.InvokeMember("YoungsModulus", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "120" });
            cstAppType.InvokeMember("PoissonsRatio", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0.33" });
            cstAppType.InvokeMember("ThermalExpansionRate", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "17" });
            cstAppType.InvokeMember("Colour", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "1", "1", "0" });
            cstAppType.InvokeMember("Wireframe", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "False" });
            cstAppType.InvokeMember("Reflection", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "False" });
            cstAppType.InvokeMember("Allowoutline", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "True" });
            cstAppType.InvokeMember("Transparentoutline", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "False" });
            cstAppType.InvokeMember("Transparency", BindingFlags.InvokeMethod, null, materialCopper, new object[] { "0" });
            cstAppType.InvokeMember("Create", BindingFlags.InvokeMethod, null, materialCopper, new object[] { });

            object Component = cstAppType.InvokeMember("Component", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("New", BindingFlags.InvokeMethod, null, Component, new object[] { "Name" });

            object Cylinder = cstAppType.InvokeMember("Cylinder", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("Reset", BindingFlags.InvokeMethod, null, Cylinder, new object[] { });
            cstAppType.InvokeMember("Name", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "vibr" });
            cstAppType.InvokeMember("Component", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "Name" });
            cstAppType.InvokeMember("Material", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "Copper (annealed)" });
            cstAppType.InvokeMember("OuterRadius", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "2" });
            cstAppType.InvokeMember("InnerRadius", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "0" });
            cstAppType.InvokeMember("Axis", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "y" });
            cstAppType.InvokeMember("Yrange", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "0", "15" });
            cstAppType.InvokeMember("Xcenter", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "0" });
            cstAppType.InvokeMember("Zcenter", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "0" });
            cstAppType.InvokeMember("Segments", BindingFlags.InvokeMethod, null, Cylinder, new object[] { "0" });
            cstAppType.InvokeMember("Create", BindingFlags.InvokeMethod, null, Cylinder, new object[] { });

            object Brick = cstAppType.InvokeMember("Brick", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("Reset", BindingFlags.InvokeMethod, null, Brick, new object[] { });
            cstAppType.InvokeMember("Name", BindingFlags.InvokeMethod, null, Brick, new object[] { "Podlojka" });
            cstAppType.InvokeMember("Component", BindingFlags.InvokeMethod, null, Brick, new object[] { "Name" });
            cstAppType.InvokeMember("Material", BindingFlags.InvokeMethod, null, Brick, new object[] { "Copper (annealed)" });
            cstAppType.InvokeMember("Xrange", BindingFlags.InvokeMethod, null, Brick, new object[] { "-15", "15" });
            cstAppType.InvokeMember("Yrange", BindingFlags.InvokeMethod, null, Brick, new object[] { "0", "5" });
            cstAppType.InvokeMember("Zrange", BindingFlags.InvokeMethod, null, Brick, new object[] { "-15", "15" });
            cstAppType.InvokeMember("Create", BindingFlags.InvokeMethod, null, Brick, new object[] { });

            cstAppType.InvokeMember("Reset", BindingFlags.InvokeMethod, null, Brick, new object[] { });
            cstAppType.InvokeMember("Name", BindingFlags.InvokeMethod, null, Brick, new object[] { "podlojkasubs" });
            cstAppType.InvokeMember("Component", BindingFlags.InvokeMethod, null, Brick, new object[] { "Name" });
            cstAppType.InvokeMember("Material", BindingFlags.InvokeMethod, null, Brick, new object[] { "Copper (annealed)" });
            cstAppType.InvokeMember("Xrange", BindingFlags.InvokeMethod, null, Brick, new object[] { "-14", "14" });
            cstAppType.InvokeMember("Yrange", BindingFlags.InvokeMethod, null, Brick, new object[] { "0", "5" });
            cstAppType.InvokeMember("Zrange", BindingFlags.InvokeMethod, null, Brick, new object[] { "-14", "14" });
            cstAppType.InvokeMember("Create", BindingFlags.InvokeMethod, null, Brick, new object[] { });

            object Solid = cstAppType.InvokeMember("Solid", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("Subtract", BindingFlags.InvokeMethod, null, Solid, new object[] { "Name" + ":Podlojka", "Name" + ":podlojkasubs" });


            /* Starting solver 
            object Solver = cstAppType.InvokeMember("Solver", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("Method", BindingFlags.InvokeMethod, null, Solver, new object[] { "Hexahedral" });
            cstAppType.InvokeMember("CalculationType", BindingFlags.InvokeMethod, null, Solver, new object[] { "TD-S" });
            cstAppType.InvokeMember("StimulationPort", BindingFlags.InvokeMethod, null, Solver, new object[] { "All" });
            cstAppType.InvokeMember("StimulationMode", BindingFlags.InvokeMethod, null, Solver, new object[] { "All" });
            cstAppType.InvokeMember("SteadyStateLimit", BindingFlags.InvokeMethod, null, Solver, new object[] { "-40" });
            cstAppType.InvokeMember("MeshAdaption", BindingFlags.InvokeMethod, null, Solver, new object[] { "False" });
            cstAppType.InvokeMember("AutoNormImpedance", BindingFlags.InvokeMethod, null, Solver, new object[] { "False" });
            cstAppType.InvokeMember("NormingImpedance", BindingFlags.InvokeMethod, null, Solver, new object[] { "50" });
            cstAppType.InvokeMember("CalculateModesOnly", BindingFlags.InvokeMethod, null, Solver, new object[] { "False" });
            cstAppType.InvokeMember("SParaSymmetry", BindingFlags.InvokeMethod, null, Solver, new object[] { "False" });
            cstAppType.InvokeMember("FullDeembedding", BindingFlags.InvokeMethod, null, Solver, new object[] { "False" });
            cstAppType.InvokeMember("SuperimposePLWExcitation", BindingFlags.InvokeMethod, null, Solver, new object[] { "False" });
            cstAppType.InvokeMember("UseSensitivityAnalysis", BindingFlags.InvokeMethod, null, Solver, new object[] { "False" });
            cstAppType.InvokeMember("Start", BindingFlags.InvokeMethod, null, Solver, new object[] { }); 
            */

            //The way to obtain S-parameter into .txt file
            //Using this example you can understand how to set another project options(project name, project saving  etc.). Use Microwave Studio documentation
            /*
            cstAppType.InvokeMember("SelectTreeItem", BindingFlags.InvokeMethod, null, cstDocument, new object[] { @"1D Results\S-Parameters\S" + Convert.ToString(i) + "," + Convert.ToString(j) + "" });
            object export = cstAppType.InvokeMember("ASCIIExport", BindingFlags.InvokeMethod, null, cstDocument, new object[] { });
            cstAppType.InvokeMember("Reset", BindingFlags.InvokeMethod, null, export, new object[] { });
            cstAppType.InvokeMember("FileName", BindingFlags.InvokeMethod, null, export, new object[] { AppDomain.CurrentDomain.BaseDirectory + "ResultS" + Convert.ToString(i) + "," + Convert.ToString(j) + ".txt" });
            cstAppType.InvokeMember("Mode", BindingFlags.InvokeMethod, null, export, new object[] { @"FixedNumber" });
            cstAppType.InvokeMember("Step", BindingFlags.InvokeMethod, null, export, new object[] { @"1001" });
            cstAppType.InvokeMember("Execute", BindingFlags.InvokeMethod, null, export, new object[] { });
            */

        }
    }
}

