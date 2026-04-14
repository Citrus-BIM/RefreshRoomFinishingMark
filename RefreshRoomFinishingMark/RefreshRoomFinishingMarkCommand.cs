using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace RefreshRoomFinishingMark
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class RefreshRoomFinishingMarkCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _ = GetPluginStartInfo();

            UIDocument? uiDoc = commandData.Application.ActiveUIDocument;
            if (uiDoc?.Document == null)
            {
                message = "Нет активного документа.";
                return Result.Cancelled;
            }

            Document doc = uiDoc.Document;
            int.TryParse(commandData.Application.Application.VersionNumber, out int versionNumber);
            if (versionNumber >= 2022)
            {
                TaskDialog.Show("Revit", "Начиная с версии Revit 2022 в ключевых спецификациях можно использовать Общие параметры! Данный плагин не нужен! Обратитесь к инструкции для получения разъяснений!");
                //return Result.Cancelled;
            }

            Room? roomFromParam = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(SpatialElement))
                .Where(e => e.GetType() == typeof(Room))
                .Cast<Room>()
                .FirstOrDefault(r => r.Area > 0);

            if (roomFromParam == null)
            {
                TaskDialog.Show("Revit", "В проекте нет размещенных помещений!");
                return Result.Cancelled;
            }

            //АР_ТипПола_Плагин
            Guid floorTypePluginParamGUID = new Guid("34eac6db-112c-49fc-a3dc-7e3470b9656f");
            Parameter floorTypePluginParam = roomFromParam.get_Parameter(floorTypePluginParamGUID);
            if (floorTypePluginParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ТипПола_Плагин\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаСтен_Плагин
            Guid wallFinishPluginParamGUID = new Guid("326001d8-4e61-494a-8eeb-74e190e11bcf");
            Parameter wallFinishPluginParam = roomFromParam.get_Parameter(wallFinishPluginParamGUID);
            if (wallFinishPluginParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаСтен_Плагин\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаСтенСнизу_Плагин
            Guid bottomWallFinishPluginParamGUID = new Guid("ee4a5627-a73d-40fa-80a9-d8702abb2f89");
            Parameter bottomWallFinishPluginParam = roomFromParam.get_Parameter(bottomWallFinishPluginParamGUID);
            if (bottomWallFinishPluginParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаСтенСнизу_Плагин\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаПотолка_Плагин
            Guid ceilingFinishPluginParamGUID = new Guid("08ea6382-373d-4347-b723-4eab2427d250");
            Parameter ceilingFinishPluginParam = roomFromParam.get_Parameter(ceilingFinishPluginParamGUID);
            if (ceilingFinishPluginParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаПотолка_Плагин\"!");
                return Result.Cancelled;
            }

            //АР_ТипПола_Ключ
            Parameter floorTypeKeyParam = roomFromParam.LookupParameter("АР_ТипПола_Ключ");
            if (floorTypeKeyParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ТипПола_Ключ\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаСтен_Ключ
            Parameter wallFinishKeyParam = roomFromParam.LookupParameter("АР_ОтделкаСтен_Ключ");
            if (wallFinishKeyParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаСтен_Ключ\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаСтенСнизу_Ключ
            Parameter bottomWallFinishKeyParam = roomFromParam.LookupParameter("АР_ОтделкаСтенСнизу_Ключ");
            if (bottomWallFinishKeyParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаСтенСнизу_Ключ\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаПотолка_Ключ
            Parameter ceilingFinishKeyParam = roomFromParam.LookupParameter("АР_ОтделкаПотолка_Ключ");
            if (ceilingFinishKeyParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаПотолка_Ключ\"!");
                return Result.Cancelled;
            }

            List<Room> roomList = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(SpatialElement))
                .Where(e => e.GetType() == typeof(Room))
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .ToList();
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Обновление марок помещений");

                foreach (Room room in roomList)
                {
                    var pFloorKey = room.LookupParameter("АР_ТипПола_Ключ");
                    long floorKeyVal = GetKeyElementIdNumeric(pFloorKey);
                    if (floorKeyVal == -1)
                        room.get_Parameter(floorTypePluginParamGUID).Set("");
                    else
                        room.get_Parameter(floorTypePluginParamGUID).Set(pFloorKey?.AsValueString() ?? "");

                    var pWallKey = room.LookupParameter("АР_ОтделкаСтен_Ключ");
                    long wallKeyVal = GetKeyElementIdNumeric(pWallKey);
                    if (wallKeyVal == -1)
                        room.get_Parameter(wallFinishPluginParamGUID).Set("");
                    else
                        room.get_Parameter(wallFinishPluginParamGUID).Set(pWallKey?.AsValueString() ?? "");

                    var pBottomWallKey = room.LookupParameter("АР_ОтделкаСтенСнизу_Ключ");
                    long bottomWallKeyVal = GetKeyElementIdNumeric(pBottomWallKey);
                    if (bottomWallKeyVal == -1)
                        room.get_Parameter(bottomWallFinishPluginParamGUID).Set("");
                    else
                        room.get_Parameter(bottomWallFinishPluginParamGUID).Set(pBottomWallKey?.AsValueString() ?? "");

                    var pCeilingKey = room.LookupParameter("АР_ОтделкаПотолка_Ключ");
                    long ceilingKeyVal = GetKeyElementIdNumeric(pCeilingKey);
                    if (ceilingKeyVal == -1)
                        room.get_Parameter(ceilingFinishPluginParamGUID).Set("");
                    else
                        room.get_Parameter(ceilingFinishPluginParamGUID).Set(pCeilingKey?.AsValueString() ?? "");
                }

                t.Commit();
            }

            TaskDialog.Show("Revit", "Обработка завершена!");
            return Result.Succeeded;
        }
        /// <summary>
        /// Числовое значение ключа спецификации: 0 если параметра нет или ElementId отсутствует; -1 для InvalidElementId; иначе IntegerValue/Value.
        /// </summary>
        private static long GetKeyElementIdNumeric(Parameter? p)
        {
            if (p == null) return 0;
#if R2019 || R2020 || R2021 || R2022 || R2023 || R2024 || R2025
            ElementId id = p.AsElementId();
            if (id == null) return 0;
            return id.IntegerValue;
#else
            return p.AsElementId().Value;
#endif
        }

        private static async Task GetPluginStartInfo()
        {
            try
            {
                Assembly thisAssembly = Assembly.GetExecutingAssembly();
                string assemblyName = "RefreshRoomFinishingMark";
                string assemblyNameRus = "Обновить марки отделки";
                string? assemblyFolderPath = Path.GetDirectoryName(thisAssembly.Location);
                if (string.IsNullOrEmpty(assemblyFolderPath)) return;

                string? parentFolder = Directory.GetParent(assemblyFolderPath)?.FullName;
                if (string.IsNullOrEmpty(parentFolder)) return;

                string dllPath = Path.Combine(parentFolder, "PluginInfoCollector", "PluginInfoCollector.dll");
                if (!File.Exists(dllPath)) return;

                Assembly assembly = Assembly.LoadFrom(dllPath);
                Type? type = assembly.GetType("PluginInfoCollector.InfoCollector");

                if (type != null)
                {
                    object? instance = Activator.CreateInstance(type);
                    var method = type.GetMethod("CollectPluginUsageAsync");

                    if (method != null && instance != null)
                    {
                        Task task = (Task)method.Invoke(instance, new object[] { assemblyName, assemblyNameRus })!;
                        await task.ConfigureAwait(false);
                    }
                }
            }
            catch
            {
                // Сбор телеметрии не должен влиять на работу команды.
            }
        }
    }
}
