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
            try
            {
                _ = GetPluginStartInfo();
            }
            catch { }

            Document doc = commandData.Application.ActiveUIDocument.Document;
            int.TryParse(commandData.Application.Application.VersionNumber, out int versionNumber);
            if (versionNumber >= 2022)
            {
                TaskDialog.Show("Revit", "Начиная с версии Revit 2022 в ключевых спецификациях можно использовать Общие параметры! Данный плагин не нужен! Обратитесь к инструкции для получения разъяснений!");
                //return Result.Cancelled;
            }

            Room roomFromParam = new FilteredElementCollector(doc)
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
                    long floorKeyVal =
#if R2019 || R2020 || R2021 || R2022 || R2023 || R2024 || R2025
                        (pFloorKey?.AsElementId() == null) ? 0 : pFloorKey.AsElementId().IntegerValue;
#else
                        (pFloorKey?.AsElementId() == null) ? 0 : pFloorKey.AsElementId().Value;
#endif
                    if (floorKeyVal == -1)
                        room.get_Parameter(floorTypePluginParamGUID).Set("");
                    else
                        room.get_Parameter(floorTypePluginParamGUID).Set(pFloorKey?.AsValueString() ?? "");

                    var pWallKey = room.LookupParameter("АР_ОтделкаСтен_Ключ");
                    long wallKeyVal =
#if R2019 || R2020 || R2021 || R2022 || R2023 || R2024 || R2025
                        (pWallKey?.AsElementId() == null) ? 0 : pWallKey.AsElementId().IntegerValue;
#else
                        (pWallKey?.AsElementId() == null) ? 0 : pWallKey.AsElementId().Value;
#endif
                    if (wallKeyVal == -1)
                        room.get_Parameter(wallFinishPluginParamGUID).Set("");
                    else
                        room.get_Parameter(wallFinishPluginParamGUID).Set(pWallKey?.AsValueString() ?? "");

                    var pBottomWallKey = room.LookupParameter("АР_ОтделкаСтенСнизу_Ключ");
                    long bottomWallKeyVal =
#if R2019 || R2020 || R2021 || R2022 || R2023 || R2024 || R2025
                        (pBottomWallKey?.AsElementId() == null) ? 0 : pBottomWallKey.AsElementId().IntegerValue;
#else
                        (pBottomWallKey?.AsElementId() == null) ? 0 : pBottomWallKey.AsElementId().Value;
#endif
                    if (bottomWallKeyVal == -1)
                        room.get_Parameter(bottomWallFinishPluginParamGUID).Set("");
                    else
                        room.get_Parameter(bottomWallFinishPluginParamGUID).Set(pBottomWallKey?.AsValueString() ?? "");

                    var pCeilingKey = room.LookupParameter("АР_ОтделкаПотолка_Ключ");
                    long ceilingKeyVal =
#if R2019 || R2020 || R2021 || R2022 || R2023 || R2024 || R2025
                        (pCeilingKey?.AsElementId() == null) ? 0 : pCeilingKey.AsElementId().IntegerValue;
#else
                        (pCeilingKey?.AsElementId() == null) ? 0 : pCeilingKey.AsElementId().Value;
#endif
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
        private static async Task GetPluginStartInfo()
        {
            // Получаем сборку, в которой выполняется текущий код
            Assembly thisAssembly = Assembly.GetExecutingAssembly();
            string assemblyName = "RefreshRoomFinishingMark";
            string assemblyNameRus = "Обновить марки отделки";
            string assemblyFolderPath = Path.GetDirectoryName(thisAssembly.Location);

            int lastBackslashIndex = assemblyFolderPath.LastIndexOf("\\");
            string dllPath = assemblyFolderPath.Substring(0, lastBackslashIndex + 1) + "PluginInfoCollector\\PluginInfoCollector.dll";

            Assembly assembly = Assembly.LoadFrom(dllPath);
            Type type = assembly.GetType("PluginInfoCollector.InfoCollector");

            if (type != null)
            {
                // Создание экземпляра класса
                object instance = Activator.CreateInstance(type);

                // Получение метода CollectPluginUsageAsync
                var method = type.GetMethod("CollectPluginUsageAsync");

                if (method != null)
                {
                    // Вызов асинхронного метода через reflection
                    Task task = (Task)method.Invoke(instance, new object[] { assemblyName, assemblyNameRus });
                    await task;  // Ожидание завершения асинхронного метода
                }
            }
        }
    }
}
