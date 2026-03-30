using System;
using System.Reflection;
using System.Linq;

class Program {
    static void Main() {
        try {
            var asm = Assembly.LoadFrom(@"c:\Users\Daryl Banks\source\repos\C3DKP26_FIXED\C3dProjects25\bin\x64\Debug\net8.0-windows\C3dProjects25_v8.dll");
            var types = asm.GetTypes();
            foreach(var t in types) {
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);
                bool hasCmd = false;
                foreach(var m in methods) {
                    if (m.CustomAttributes.Any(a => a.AttributeType.Name == "CommandMethodAttribute")) {
                        hasCmd = true;
                        break;
                    }
                }
                if (hasCmd) {
                    Console.WriteLine($"[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof({t.FullName}))]");
                }
            }
        } catch (Exception ex) {
            Console.WriteLine(ex.ToString());
        }
    }
}
