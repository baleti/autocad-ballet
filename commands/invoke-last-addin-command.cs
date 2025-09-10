using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutocadBallet.InvokeLastAddinCommand))]

namespace AutocadBallet
{
    public class InvokeLastAddinCommand
    {
    private const string FolderName = "autocad-ballet";
    private const string ConfigFileName = "InvokeAddinCommand-last-dll-path";
    private const string LastCommandFileName = "InvokeAddinCommand-history";
    private static readonly string ConfigFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName);
    private static readonly string ConfigFilePath = Path.Combine(ConfigFolderPath, ConfigFileName);
    private static readonly string LastCommandFilePath = Path.Combine(ConfigFolderPath, LastCommandFileName);

    private Dictionary<string, Assembly> loadedAssemblies = new Dictionary<string, Assembly>();

    [CommandMethod("invoke-last-addin-command")]
    public void Execute()
    {
        try
        {
            if (!File.Exists(ConfigFilePath) || !File.Exists(LastCommandFilePath))
            {
                AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nNo previous command found. Run a command using InvokeAddinCommand first.");
                return;
            }

            string dllPath = File.ReadAllText(ConfigFilePath);
            
            string commandClassName = GetLastCommand();
            
            if (string.IsNullOrEmpty(commandClassName))
            {
                AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nNo command history found.");
                return;
            }

            if (string.IsNullOrEmpty(dllPath) || !File.Exists(dllPath))
            {
                AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nInvalid DLL path.");
                return;
            }

            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

            Assembly assembly = LoadAssembly(dllPath);
            Type commandType = assembly.GetType(commandClassName);

            if (commandType == null)
            {
                AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nCould not find type: {commandClassName}");
                return;
            }

            object commandInstance = Activator.CreateInstance(commandType);
            MethodInfo method = commandType.GetMethod("Execute");

            if (method != null)
            {
                method.Invoke(commandInstance, null);
            }
            else
            {
                AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nExecute method not found.");
            }
        }
        catch (System.Exception ex)
        {
            string message = $"An error occurred: {ex.Message}";
            if (ex.InnerException != null)
            {
                message += $"\nInner Exception: {ex.InnerException.Message}";
            }
            AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n{message}");
        }
        finally
        {
            AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomain_AssemblyResolve;
        }
    }

    private string GetLastCommand()
    {
        try
        {
            string[] lines = File.ReadAllLines(LastCommandFilePath);
            
            for (int i = lines.Length - 1; i >= 0; i--)
            {
                if (!string.IsNullOrWhiteSpace(lines[i]))
                {
                    return lines[i].Trim();
                }
            }
            
            return null;
        }
        catch
        {
            return null;
        }
    }

    private Assembly LoadAssembly(string assemblyPath)
    {
        string assemblyName = Path.GetFileNameWithoutExtension(assemblyPath);

        if (loadedAssemblies.ContainsKey(assemblyName))
        {
            return loadedAssemblies[assemblyName];
        }

        byte[] assemblyBytes = File.ReadAllBytes(assemblyPath);
        Assembly assembly = Assembly.Load(assemblyBytes);
        loadedAssemblies[assemblyName] = assembly;

        string directory = Path.GetDirectoryName(assemblyPath);
        foreach (string dllFile in Directory.GetFiles(directory, "*.dll"))
        {
            if (dllFile != assemblyPath)
            {
                string dllName = Path.GetFileNameWithoutExtension(dllFile);
                if (!loadedAssemblies.ContainsKey(dllName))
                {
                    try
                    {
                        byte[] dllBytes = File.ReadAllBytes(dllFile);
                        Assembly dllAssembly = Assembly.Load(dllBytes);
                        loadedAssemblies[dllName] = dllAssembly;
                    }
                    catch (BadImageFormatException)
                    {
                        continue;
                    }
                }
            }
        }

        return assembly;
    }

    private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
    {
        AssemblyName assemblyName = new AssemblyName(args.Name);
        string shortName = assemblyName.Name;

        if (loadedAssemblies.ContainsKey(shortName))
        {
            return loadedAssemblies[shortName];
        }

        if (!File.Exists(ConfigFilePath))
        {
            return null;
        }

        string dllPath = File.ReadAllText(ConfigFilePath);
        string directory = Path.GetDirectoryName(dllPath);
        string assemblyPath = Path.Combine(directory, shortName + ".dll");

        if (File.Exists(assemblyPath))
        {
            try
            {
                return LoadAssembly(assemblyPath);
            }
            catch (System.Exception)
            {
                return null;
            }
        }

        return null;
    }
    }
}