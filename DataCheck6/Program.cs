using System.Text;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Utf8StringInterpolation;
using ZLogger;
using ZLogger.Providers;


//==
var builder = ConsoleApp.CreateBuilder(args);
builder.ConfigureServices((ctx,services) =>
{
    // Register appconfig.json to IOption<MyConfig>
    services.Configure<MyConfig>(ctx.Configuration);

    // Using Cysharp/ZLogger for logging to file
    services.AddLogging(logging =>
    {
        logging.ClearProviders();
        logging.SetMinimumLevel(LogLevel.Trace);
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        var utcTimeZoneInfo = TimeZoneInfo.Utc;
        logging.AddZLoggerConsole(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });
        });
        logging.AddZLoggerRollingFile(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });

            // File name determined by parameters to be rotated
            options.FilePathSelector = (timestamp, sequenceNumber) => $"logs/{timestamp.ToLocalTime():yyyy-MM-dd}_{sequenceNumber:00}.log";
            
            // The period of time for which you want to rotate files at time intervals.
            options.RollingInterval = RollingInterval.Day;
            
            // Limit of size if you want to rotate by file size. (KB)
            options.RollingSizeKB = 1024;        
        });
    });
});

var app = builder.Build();
app.AddCommands<DataCheckApp>();
app.Run();


public class DataCheckApp : ConsoleAppBase
{
    bool isAllPass = true;
    const int HOST_NAME_COLUMN = 1;
    const int USED_PORT_COLUMN = 2;

    readonly ILogger<DataCheckApp> logger;
    readonly IOptions<MyConfig> config;

    Dictionary<string, List<string>> MyModelAndPortName = new Dictionary<string, List<string>>();
    Dictionary<string, List<string>> MyHostNameUsedPorts = new Dictionary<string, List<string>>();

    List<MyDevicePort> MyDevicePorts = new List<MyDevicePort>();

    public DataCheckApp(ILogger<DataCheckApp> logger,IOptions<MyConfig> config)
    {
        this.logger = logger;
        this.config = config;
    }

//    [Command("")]
    public void Check(string diagram, string excelpath, string prefix, int devicelength, int rosettelength)
    {
//== start
        logger.ZLogInformation($"==== tool {getMyFileVersion()} ====");
        
        if (!File.Exists(diagram))
        {
            logger.ZLogError($"[NG] target diagrm file is missing.");
            return;
        }
        if (!File.Exists(excelpath))
        {
            logger.ZLogError($"[NG] target excel file is missing.");
            return;
        }

        logger.ZLogInformation($"== パラメーター ==");
        logger.ZLogInformation($"対象:{diagram}");
        try
        {
            using FileStream fsDiagram = new FileStream(diagram, FileMode.Open, FileAccess.Read, FileShare.Read);
            using XLWorkbook xlWorkbookDiagram = new XLWorkbook(fsDiagram);
            IXLWorksheets sheetsDiagram = xlWorkbookDiagram.Worksheets;
            foreach (IXLWorksheet? sheet in sheetsDiagram)
            {
                if (sheet != null && sheet.Name == "diagram")
                {
                    int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                    logger.ZLogInformation($"ネットワーク構成図:{sheet.Cell(1, 2).Value}, 最後の行:{lastUsedRowNumber}");
                    string tmpHostname = "";
                    for (int r = 3; r < lastUsedRowNumber + 1; r++)
                    {
                        List<string> tmpUsedPorts = new List<string>();
                        tmpHostname = sheet.Cell(r, HOST_NAME_COLUMN).Value.ToString();
                        string usedPorts = sheet.Cell(r, USED_PORT_COLUMN).Value.ToString();
                        foreach (var port in usedPorts.Split(";"))
                        {
                            if (!string.IsNullOrEmpty(port))
                            {
                                tmpUsedPorts.Add(port.TrimStart().TrimEnd());
                            }
                        }
                        MyHostNameUsedPorts.Add(tmpHostname, tmpUsedPorts);                
                    }
                }
                else
                {
                    isAllPass = false;
                    logger.ZLogError($"[NG] {diagram}ファイル内にシート名：diagramが見つかりませんでした");
                }
            }
        }
        catch (IOException ie)
        {
            logger.ZLogError($"[ERROR] Excelファイルの読み取りでエラーが発生しました。Excelで対象ファイルを開いていませんか？ 詳細:({ie.Message})");
            return;
        }
        catch (System.Exception)
        {
            throw;
        }

        try
        {
            using FileStream fs = new FileStream(excelpath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using XLWorkbook xlWorkbook = new XLWorkbook(fs);
            IXLWorksheets sheets = xlWorkbook.Worksheets;

//== init
            int deviceFromCableIdColumn = config.Value.DeviceFromCableIdColumn;
            int deviceFromConnectColumn = config.Value.DeviceFromConnectColumn;
            int deviceFromFloorNameColumn = config.Value.DeviceFromFloorNameColumn;
            int deviceFromDeviceNameColumn = config.Value.DeviceFromDeviceNameColumn;
            int deviceFromDeviceNumberColumn = config.Value.DeviceFromDeviceNumberColumn;
            int deviceFromHostNameColumn = config.Value.DeviceFromHostNameColumn;
            int deviceFromModelNameColumn = config.Value.DeviceFromModelNameColumn;
            int deviceFromPortNameColumn = config.Value.DeviceFromPortNameColumn;
            int deviceFromConnectorNameColumn = config.Value.DeviceFromConnectorNameColumn;
            int deviceToFloorNameColumn = config.Value.DeviceToFloorNameColumn;
            int deviceToDeviceNameColumn = config.Value.DeviceToDeviceNameColumn;
            int deviceToDeviceNumberColumn = config.Value.DeviceToDeviceNumberColumn;
            int deviceToModelNameColumn = config.Value.DeviceToModelNameColumn;
            int deviceToHostNameColumn = config.Value.DeviceToHostNameColumn;
            int deviceToPortNameColumn = config.Value.DeviceToPortNameColumn;
            string wordConnect = config.Value.WordConnect;
            string wordDisconnect = config.Value.WordDisconnect;
            int hostNameLength = devicelength;
            string ignoreDeviceNameToHostNameLength = config.Value.IgnoreDeviceNameToHostNameLength;
            int rosetteHostNameLength = rosettelength;
            string ignoreDeviceNameToHostNamePrefix = config.Value.IgnoreDeviceNameToHostNamePrefix;
            string ignoreDeviceNameToConnectXConnect = config.Value.IgnoreDeviceNameToConnectXConnect;
            string ignoreConnectorNameToAll = config.Value.IgnoreConnectorNameToAll;
            string wordDeviceToHostNameList = config.Value.WordDeviceToHostNameList;
            string deviceNameToRosette = config.Value.DeviceNameToRosette;

            logger.ZLogInformation($"== パラメーター ==");
            logger.ZLogInformation($"対象:{excelpath}");
            logger.ZLogInformation($"接頭語:{prefix}, ホスト名の長さ:{hostNameLength}, ローゼット名の長さ:{rosetteHostNameLength}");

            foreach (IXLWorksheet? sheet in sheets)
            {
                if (sheet != null)
                {
                    int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
                    int lastUsedColumNumber = sheet.LastColumnUsed() == null ? 0 : sheet.LastColumnUsed().ColumnNumber();

                    logger.ZLogInformation($"シート名:{sheet.Name}, 最後の行:{lastUsedRowNumber}, 最後の列:{lastUsedColumNumber}");

                    for (int r = 1; r < lastUsedRowNumber + 1; r++)
                    {
                        IXLCell cellConnect = sheet.Cell(r, deviceFromConnectColumn);
                        IXLCell cellCableID = sheet.Cell(r, deviceFromCableIdColumn);
                        if (cellConnect.IsEmpty() == true)
                        {
                            // nothing
                        }
                        else
                        {
                            if (cellConnect.Value.GetText() == wordConnect || cellConnect.Value.GetText() == wordDisconnect)
                            {
                                MyDevicePort tmpDevicePort = new MyDevicePort();
                                tmpDevicePort.fromConnect = cellConnect.Value.GetText();
                                int id = -1;
                                switch (cellCableID.DataType)
                                {
                                    case XLDataType.Number:
                                        id = cellCableID.GetValue<int>();
                                        break;
                                    case XLDataType.Text:
                                        try
                                        {
                                            id = int.Parse(cellCableID.GetValue<string>());
                                        }
                                        catch (System.FormatException)
                                        {
                                            logger.ZLogError($"ID is NOT type ( Text-> parse ) at sheet:{sheet.Name} row:{r}");
                                            continue;
                                        }
                                        catch (System.Exception)
                                        {
                                            throw;
                                        }
                                        break;
                                    default:
                                        logger.ZLogError($"ID is NOT type ( Number | Text ) at sheet:{sheet.Name} row:{r}");
                                        continue;
                                }
                                tmpDevicePort.fromCableID = id;
                                tmpDevicePort.fromFloorName = sheet.Cell(r, deviceFromFloorNameColumn).Value.ToString();
                                tmpDevicePort.fromDeviceName = sheet.Cell(r, deviceFromDeviceNameColumn).Value.ToString();
                                tmpDevicePort.fromDeviceNumber = sheet.Cell(r, deviceFromDeviceNumberColumn).Value.ToString();
                                tmpDevicePort.fromHostName = sheet.Cell(r, deviceFromHostNameColumn).Value.ToString();
                                tmpDevicePort.fromModelName = sheet.Cell(r, deviceFromModelNameColumn).Value.ToString();
                                tmpDevicePort.fromPortName = sheet.Cell(r, deviceFromPortNameColumn).Value.ToString();
                                tmpDevicePort.fromConnectorName = sheet.Cell(r, deviceFromConnectorNameColumn).Value.ToString();
                                tmpDevicePort.toFloorName = sheet.Cell(r, deviceToFloorNameColumn).Value.ToString();
                                tmpDevicePort.toDeviceName = sheet.Cell(r, deviceToDeviceNameColumn).Value.ToString();
                                tmpDevicePort.toDeviceNumber = sheet.Cell(r, deviceToDeviceNumberColumn).Value.ToString();
                                tmpDevicePort.toModelName = sheet.Cell(r, deviceToModelNameColumn).Value.ToString();
                                tmpDevicePort.toHostName = sheet.Cell(r, deviceToHostNameColumn).Value.ToString();
                                tmpDevicePort.toPortName = sheet.Cell(r, deviceToPortNameColumn).Value.ToString();
                                MyDevicePorts.Add(tmpDevicePort);
                            }
                        }
                    }
                }
            }

//== create ModelAndPortName
            readModeAndPortName();

//== print
            printMyHostNameUsedPorts();
            printMyDevicePorts();

//== check CableList ModeAndPortName
            checkModelAndPortName();

//== Diagram VS CableList
            checkDiagramVsCableList();

//== check duplicate CableID
            checkDuplicateCableId();

//== check toDevice --> Connect
            checkToDeviceAtFromConnect();

//== check hostname count
            checkHostNameLength(hostNameLength, rosetteHostNameLength);

//== check hostname prefix
            checkHostNamePrefix(prefix);

//== check device word to hostName word
            checkDeviceToHostName(prefix);

//== check device&number to hostName
            checkDeviceAndNumberToHostName();

//== check rosette
            checkRosette();

//== check 
            checkConnectXConnect();
        }
        catch (IOException ie)
        {
            logger.ZLogError($"[ERROR] Excelファイルの読み取りでエラーが発生しました。Excelで対象ファイルを開いていませんか？ 詳細:({ie.Message})");
            return;
        }
        catch (System.Exception)
        {
            throw;
        }

//== finish
        if (isAllPass)
        {
            logger.ZLogInformation($"== [Congratulations!] すべての確認項目をパスしました ==");
        }
        logger.ZLogInformation($"==== tool finish ====");

    }

    private void readModeAndPortName()
    {
        logger.ZLogTrace($"== start readModeAndPortName ==");
        string modelAndPortName = config.Value.ModelAndPortName;
        foreach (var modelport in modelAndPortName.Split(','))
        {
            List<string> port = new List<string>();
            string[] item = modelport.Split('|');
            string model = item[0];
            foreach (var portname in item[1].Split(';'))
            {
                if (string.IsNullOrEmpty(portname))
                {
                    logger.ZLogDebug($"ポート名が、NULLまたは空白でした");
                    continue;
                }
                else
                {
                    port.Add(portname);
                }
            }
            if (MyModelAndPortName.ContainsKey(model))
            {
                logger.ZLogDebug($"既にモデル名が登録されていました");
                continue;
            }
            else
            {
                MyModelAndPortName.Add(model, port);
            }
        }

        logger.ZLogTrace($"== end readModeAndPortName ==");
    }

    private void checkModelAndPortName()
    {
        logger.ZLogInformation($"== start モデル名とポート名の確認 ==");
        bool isError = false;
        Dictionary<string,string> dicIgnoreModelName = new Dictionary<string, string>();
        string ignoreModelName = config.Value.IgnoreModelName;
        foreach (var ignore in ignoreModelName.Split(','))
        {
            dicIgnoreModelName.Add(ignore, "");
        }

        foreach (var device in MyDevicePorts)
        {
            // from
            if (MyModelAndPortName.ContainsKey(device.fromModelName))
            {
                var portnames = MyModelAndPortName[device.fromModelName];
                if (portnames.Contains(device.fromPortName))
                {
                    // OK
                    logger.ZLogTrace($"モデル名とポート名 check OK");
                }
                else
                {
                    isError = true;
                    logger.ZLogError($"ポート名の間違いが発見されました ケーブルID:{device.fromCableID} From側ポート名:{device.fromPortName}");
                }
            }
            else
            {
                isError = true;
                logger.ZLogError($"モデル名が存在しませんでした ケーブルID:{device.fromCableID} From側モデル名:{device.fromModelName}");
            }

            // to
            if (isNotIgnoreDevice(device.toModelName, dicIgnoreModelName))
            {
                if (MyModelAndPortName.ContainsKey(device.toModelName))
                {
                    var portnames = MyModelAndPortName[device.toModelName];
                    if (portnames.Contains(device.toPortName))
                    {
                        // OK
                        logger.ZLogTrace($"モデル名とポート名 check OK");
                    }
                    else
                    {
                        isError = true;
                        logger.ZLogError($"ポート名の間違いが発見されました ケーブルID:{device.fromCableID} To側ポート名:{device.toPortName}");
                    }
                }
                else
                {
                    isError = true;
                    logger.ZLogError($"モデル名が存在しませんでした ケーブルID:{device.fromCableID} To側モデル名:{device.toModelName}");
                }
            }
        }

        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] モデル名とポート名で、不一致が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] モデル名とポート名は正しいことが確認されました");
        }
        logger.ZLogInformation($"== end モデル名とポート名の確認 ==");
    }

    private void printMyHostNameUsedPorts()
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var hostname in MyHostNameUsedPorts.Keys)
        {
            logger.ZLogTrace($"ホスト名:{hostname},使用済ポート:{string.Join(";",MyHostNameUsedPorts[hostname])}");
        }
        logger.ZLogTrace($"== end print ==");
    }

    private void printMyDevicePorts()
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var device in MyDevicePorts)
        {
            logger.ZLogTrace($"CableID:{device.fromCableID},connect:{device.fromConnect},(from) Device:{device.fromDeviceName},Host:{device.fromHostName},Model:{device.toModelName},Port:{device.fromPortName},(to) Device:{device.toDeviceName},Host:{device.toHostName},Model:{device.toModelName},Port:{device.toPortName}");
        }
        logger.ZLogTrace($"== end print ==");
    }

    private void checkDuplicateCableId()
    {
        logger.ZLogInformation($"== start ケーブルIDの重複の確認 ==");
        bool isError = false;
        Dictionary<int, string> cableId = new Dictionary<int, string>();
        foreach (var device in MyDevicePorts)
        {
            try
            {
                cableId.Add(device.fromCableID, device.fromHostName +"&"+ device.fromPortName);
            }
            catch (System.ArgumentException)
            {
                isError = true;
                logger.ZLogError($"重複エラー ケーブルID:{device.fromCableID} ( {cableId[device.fromCableID]} | {device.fromHostName}&{device.fromPortName} )");
            }
            catch (System.Exception)
            {
                throw;
            }
        }
        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] ケーブルIDの重複が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] ケーブルIDの重複はありませんでした");
        }
        logger.ZLogInformation($"== end ケーブルIDの重複の確認 ==");
    }

    private bool isNotIgnoreDevice(string device, Dictionary<string,string> dicIgnore)
    {
        return !dicIgnore.ContainsKey(device);
    }

    private bool isDevice(string device, Dictionary<string,string> dicDevice)
    {
        return dicDevice.ContainsKey(device);
    }

    private void checkHostNameLength(int deviceHostNameLength, int rosetteHostNameLength)
    {
        logger.ZLogInformation($"== start ホスト名の長さの確認 ==");
        bool isError = false;
        Dictionary<string,string> dicIgnoreDeviceName = new Dictionary<string, string>();
        string ignoreDeviceNameToHostNameLength = config.Value.IgnoreDeviceNameToHostNameLength;
        Dictionary<string,string> dicIgnoreConnectorName = new Dictionary<string, string>();
        string ignoreConnectorNameToAll = config.Value.IgnoreConnectorNameToAll;
        foreach (var ignore in ignoreDeviceNameToHostNameLength.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }
        foreach (var ignore in ignoreConnectorNameToAll.Split(','))
        {
            dicIgnoreConnectorName.Add(ignore, "");
        }

        string wordConnect = config.Value.WordConnect;
        foreach (var device in MyDevicePorts)
        {
            if (device.fromHostName.Length != deviceHostNameLength)
            {
                if (isNotIgnoreDevice(device.fromDeviceName, dicIgnoreDeviceName) && isNotIgnoreDevice(device.fromConnectorName, dicIgnoreConnectorName))
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} From側デバイス名:{device.fromDeviceName} From側ホスト名:{device.fromHostName}");
                }
                else
                {
                    logger.ZLogTrace($"[checkHostNameLength] 除外しました ケーブルID:{device.fromCableID} From側デバイス名:{device.fromDeviceName} From側コネクター形状:{device.fromConnectorName}");
                }
            }
            else
            {
                logger.ZLogTrace($"[checkHostNameLength] 文字数({deviceHostNameLength})で一致しました ケーブルID:{device.fromCableID} From側ホスト名:{device.fromHostName}");
            }

            if (device.fromConnect == wordConnect)
            {
                if (device.toHostName.Length != deviceHostNameLength)
                {
                    if (device.toHostName.Length != rosetteHostNameLength)
                    {
                        if (isNotIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName) && isNotIgnoreDevice(device.fromConnectorName, dicIgnoreConnectorName))
                        {
                            isError = true;
                            logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName} To側ホスト名:{device.toHostName}");
                        }
                        else
                        {
                            logger.ZLogTrace($"[checkHostNameLength] 除外しました ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName} From側コネクター形状:{device.fromConnectorName}");
                        }
                    }
                }
                else
                {
                    logger.ZLogTrace($"[checkHostNameLength] 文字数({deviceHostNameLength})で一致しました ケーブルID:{device.fromCableID} To側ホスト名:{device.toHostName}");
                }
            }
        }
        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] ホスト名の長さの不一致が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] ホスト名の長さの不一致はありませんでした");
        }
        logger.ZLogInformation($"== end ホスト名の長さの確認 ==");
    }

    private void checkHostNamePrefix(string prefix)
    {
        logger.ZLogInformation($"== start ホスト名の接頭語の確認 ==");
        bool isError = false;
        Dictionary<string,string> dicIgnoreDeviceName = new Dictionary<string, string>();
        string ignoreDeviceNameToHostNamePrefix = config.Value.IgnoreDeviceNameToHostNamePrefix;
        Dictionary<string,string> dicIgnoreConnectorName = new Dictionary<string, string>();
        string ignoreConnectorNameToAll = config.Value.IgnoreConnectorNameToAll;
        foreach (var ignore in ignoreDeviceNameToHostNamePrefix.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }
        foreach (var ignore in ignoreConnectorNameToAll.Split(','))
        {
            dicIgnoreConnectorName.Add(ignore, "");
        }

        string wordConnect = config.Value.WordConnect;
        foreach (var device in MyDevicePorts)
        {
            if (!device.fromHostName.StartsWith(prefix))
            {
                if (isNotIgnoreDevice(device.fromDeviceName, dicIgnoreDeviceName) && isNotIgnoreDevice(device.fromConnectorName, dicIgnoreConnectorName))
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} 接頭語:{prefix} From側ホスト名:{device.fromHostName}");
                }
                else
                {
                    logger.ZLogTrace($"[checkHostNamePrefix] 除外しました ケーブルID:{device.fromCableID} From側デバイス名:{device.fromDeviceName} From側コネクター形状:{device.fromConnectorName}");
                }
            }
            else
            {
                logger.ZLogTrace($"[checkHostNamePrefix] 接頭語({prefix})で一致しました ケーブルID:{device.fromCableID} From側ホスト名:{device.fromHostName}");
            }

            if (device.fromConnect == wordConnect)
            {
                if (!device.toHostName.StartsWith(prefix))
                {
                    if (isNotIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName) && isNotIgnoreDevice(device.fromConnectorName, dicIgnoreConnectorName))
                    {
                        isError = true;
                        logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} 接頭語:{prefix} To側ホスト名:{device.toHostName}");
                    }
                    else
                    {
                        logger.ZLogTrace($"[checkHostNamePrefix] 除外しました ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName} From側コネクター形状:{device.fromConnectorName}");
                    }
                }
                else
                {
                    logger.ZLogTrace($"[checkHostNamePrefix] 接頭語({prefix})で一致しました ケーブルID:{device.fromCableID} To側ホスト名:{device.toHostName}");
                }
            }
        }
        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] ホスト名の接頭語の不一致が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] ホスト名の接頭語の不一致はありませんでした");
        }
        logger.ZLogInformation($"== end ホスト名の接頭語の確認 ==");
    }

    private void checkDeviceToHostName(string prefix)
    {
        logger.ZLogInformation($"== start 装置名によるホスト名に含む文字列の確認 ==");
        bool isError = false;
        Dictionary<string,string> dicIgnoreDeviceName = new Dictionary<string, string>();
        string ignoreDeviceNameToHostNamePrefix = config.Value.IgnoreDeviceNameToHostNamePrefix;
        Dictionary<string,string> dicIgnoreConnectorName = new Dictionary<string, string>();
        string ignoreConnectorNameToAll = config.Value.IgnoreConnectorNameToAll;
        Dictionary<string,string> dicDeviceToHostName = new Dictionary<string, string>();
        string wordDeviceToHostNameList = config.Value.WordDeviceToHostNameList;
        foreach (var ignore in ignoreDeviceNameToHostNamePrefix.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }
        foreach (var ignore in ignoreConnectorNameToAll.Split(','))
        {
            dicIgnoreConnectorName.Add(ignore, "");
        }
        foreach (var keyAndValue in wordDeviceToHostNameList.Split(','))
        {
            string[] item = keyAndValue.Split('/');
            dicDeviceToHostName.Add(item[0], item[1]);
        }

        string wordConnect = config.Value.WordConnect;
        foreach (var device in MyDevicePorts)
        {
            if (dicDeviceToHostName.ContainsKey(device.fromDeviceName))
            {
                string hostname = device.fromHostName;
                string targetHostname = hostname.Replace(prefix, "");
                if (!targetHostname.Contains(dicDeviceToHostName[device.fromDeviceName]))
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} ホスト名に含む文字列:{dicDeviceToHostName[device.fromDeviceName]} From側ホスト名:{device.fromHostName}");
                }
                else
                {
                    logger.ZLogTrace($"[checkDeviceToHostName] ホスト名に含む文字列:{dicDeviceToHostName[device.fromDeviceName]}が正しく含まれています ケーブルID:{device.fromCableID} From側ホスト名:{device.fromHostName}");
                }
            }
            else
            {
                isError = true;
                logger.ZLogError($"定義エラー ケーブルID:{device.fromCableID} Fromデバイス名:{device.fromDeviceName}に対応するホスト名に含む文字列が定義されていません");
            }

            if (device.fromConnect == wordConnect)
            {
                if (isNotIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName) && isNotIgnoreDevice(device.fromConnectorName, dicIgnoreConnectorName))
                {
                    if (dicDeviceToHostName.ContainsKey(device.toDeviceName))
                    {
                        string hostname = device.toHostName;
                        string targetHostname = hostname.Replace(prefix, "");
                        if (!targetHostname.Contains(dicDeviceToHostName[device.toDeviceName]))
                        {
                            isError = true;
                            logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} ホスト名に含む文字列:{dicDeviceToHostName[device.toDeviceName]} To側ホスト名:{device.toHostName}");
                        }
                        else
                        {
                            logger.ZLogTrace($"[checkDeviceToHostName] ホスト名に含む文字列:{dicDeviceToHostName[device.toDeviceName]}が正しく含まれています ケーブルID:{device.fromCableID} To側ホスト名:{device.toHostName}");
                        }
                    }
                    else
                    {
                        isError = true;
                        logger.ZLogError($"定義エラー ケーブルID:{device.fromCableID} Toデバイス名:{device.toDeviceName}に対応するホスト名に含む文字列が定義されていません");
                    }
                }
                else
                {
                    logger.ZLogTrace($"[checkDeviceToHostName] 除外しました ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName} From側コネクター形状:{device.fromConnectorName}");
                }
            }
        }

        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] 装置名によるホスト名に含む文字列の不一致が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] 装置名によるホスト名に含む文字列の不一致はありませんでした");
        }
        logger.ZLogInformation($"== end 装置名によるホスト名に含む文字列の確認 ==");
    }

    void checkDeviceAndNumberToHostName()
    {
        logger.ZLogInformation($"== start 装置名＆識別名とホスト名の一意の確認 ==");
        bool isError = false;
        Dictionary<string,string> dicIgnoreDeviceName = new Dictionary<string, string>();
        string ignoreDeviceNameToHostNamePrefix = config.Value.IgnoreDeviceNameToHostNamePrefix;
        Dictionary<string,string> dicIgnoreConnectorName = new Dictionary<string, string>();
        string ignoreConnectorNameToAll = config.Value.IgnoreConnectorNameToAll;
        foreach (var ignore in ignoreDeviceNameToHostNamePrefix.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }
        foreach (var ignore in ignoreConnectorNameToAll.Split(','))
        {
            dicIgnoreConnectorName.Add(ignore, "");
        }

        Dictionary<string,string> dicDeviceAndNumberToHostName = new Dictionary<string, string>();
        Dictionary<string,string> dicHostNameToDeviceAndNumber = new Dictionary<string, string>();

        string wordConnect = config.Value.WordConnect;
        foreach (var device in MyDevicePorts)
        {
            try
            {
                dicDeviceAndNumberToHostName.Add(device.fromDeviceName + device.fromDeviceNumber, device.fromHostName);
            }
            catch (System.ArgumentException)
            {
                string tmp = dicDeviceAndNumberToHostName[device.fromDeviceName + device.fromDeviceNumber];
                if (!tmp.Equals(device.fromHostName))
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} Fromデバイス名:{device.fromDeviceName + device.fromDeviceNumber} Fromホスト名:{device.fromHostName}");
                }
            }
            catch (System.Exception)
            {
                throw;
            }

            try
            {
                dicHostNameToDeviceAndNumber.Add(device.fromHostName, device.fromDeviceName + device.fromDeviceNumber);
            }
            catch (System.ArgumentException)
            {
                string tmpDevice = device.fromDeviceName + device.fromDeviceNumber;
                string tmp = dicHostNameToDeviceAndNumber[device.fromHostName];
                if (!tmp.Equals(tmpDevice))
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} Fromデバイス名:{device.fromDeviceName + device.fromDeviceNumber} Fromホスト名:{device.fromHostName}");
                }
            }
            catch (System.Exception)
            {
                throw;
            }

            if (device.fromConnect == wordConnect)
            {
                if (isNotIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName) && isNotIgnoreDevice(device.fromConnectorName, dicIgnoreConnectorName))
                {
                    try
                    {
                        dicDeviceAndNumberToHostName.Add(device.toDeviceName + device.toDeviceNumber, device.toHostName);
                    }
                    catch (System.ArgumentException)
                    {
                        string tmp = dicDeviceAndNumberToHostName[device.toDeviceName + device.toDeviceNumber];
                        if (!tmp.Equals(device.toHostName))
                        {
                            isError = true;
                            logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} Toデバイス名:{device.toDeviceName + device.toDeviceNumber} Toホスト名:{device.toHostName}");
                        }
                    }
                    catch (System.Exception)
                    {
                        throw;
                    }

                    try
                    {
                        dicHostNameToDeviceAndNumber.Add(device.toHostName, device.toDeviceName + device.toDeviceNumber);
                    }
                    catch (System.ArgumentException)
                    {
                        string tmpDevice = device.toDeviceName + device.toDeviceNumber;
                        string tmp = dicHostNameToDeviceAndNumber[device.toHostName];
                        if (!tmp.Equals(tmpDevice))
                        {
                            isError = true;
                            logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} Toデバイス名:{device.toDeviceName + device.toDeviceNumber} Toホスト名:{device.toHostName}");
                        }
                    }
                    catch (System.Exception)
                    {
                        throw;
                    }
                }
                else
                {
                    logger.ZLogTrace($"[checkDeviceAndNumberToHostName] 除外しました ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName} From側コネクター形状:{device.fromConnectorName}");
                }
            }
        }

        foreach (var key in dicDeviceAndNumberToHostName.Keys)
        {
            if (dicHostNameToDeviceAndNumber.ContainsKey(dicDeviceAndNumberToHostName[key]))
            {
                string deviceAndNumber = dicHostNameToDeviceAndNumber[dicDeviceAndNumberToHostName[key]];
                if (key.Equals(deviceAndNumber))
                {
                    logger.ZLogTrace($"[checkDeviceAndNumberToHostName] 一致しました デバイス名:{deviceAndNumber} ホスト名:{dicDeviceAndNumberToHostName[key]}");
                }
                else
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー デバイス名:{deviceAndNumber} ホスト名:{dicDeviceAndNumberToHostName[key]}");
                }
            }
            else
            {
                isError = true;
                logger.ZLogError($"キーエラー キー({key})が存在しません");
            }
        }
        foreach (var key in dicHostNameToDeviceAndNumber.Keys)
        {
            if (dicDeviceAndNumberToHostName.ContainsKey(dicHostNameToDeviceAndNumber[key]))
            {
                string hostname = dicDeviceAndNumberToHostName[dicHostNameToDeviceAndNumber[key]];
                if (key.Equals(hostname))
                {
                    logger.ZLogTrace($"[checkDeviceAndNumberToHostName] 一致しました ホスト名:{hostname} デバイス名:{dicHostNameToDeviceAndNumber[key]}");
                }
                else
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー ホスト名:{hostname} デバイス名:{dicHostNameToDeviceAndNumber[key]}");
                }
            }
            else
            {
                isError = true;
                logger.ZLogError($"キーエラー キー({key})が存在しません");
            }
        }

        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] 装置名＆識別名とホスト名の重複が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] 装置名＆識別名とホスト名は一意でした");
        }
        logger.ZLogInformation($"== end 装置名＆識別名とホスト名の一意の確認 ==");
    }

    void checkDiagramVsCableList()
    {
        logger.ZLogInformation($"== start ネットワーク構成図とケーブルリストの接続ポートの一致の確認 ==");
        bool isError = false;
        string wordConnect = config.Value.WordConnect;
        Dictionary<string,List<string>> dicFromCableList = new Dictionary<string, List<string>>();
        foreach (var device in MyDevicePorts)
        {
            if (device.fromConnect.Equals(wordConnect))
            {
                try
                {
                    dicFromCableList.Add(device.fromHostName, new List<string>());
                }
                catch (System.ArgumentException)
                {
                    logger.ZLogTrace($"nothing {device.fromHostName}");
                }
                catch (System.Exception)
                {
                    throw;
                }
                dicFromCableList[device.fromHostName].Add(device.fromPortName);
            }
            else
            {
                logger.ZLogTrace($"[checkDiagramVsCableList] ケーブルID:{device.fromCableID} は({wordConnect})以外のため確認しない");
            }
        }

        var listCableKeys = dicFromCableList.Keys;
        var listDiagramKeys = MyHostNameUsedPorts.Keys;
        var diffs = listCableKeys.Except(listDiagramKeys).Union(listDiagramKeys.Except(listCableKeys));
        if (diffs.Count() > 0)
        {
            isError = true;
            logger.ZLogError($"ケーブルリストとネットワーク構成にFromホスト名({string.Join(",",diffs)})の差分が発見されました");                
        }
        else
        {
            logger.ZLogTrace($"[checkDiagramVsCableList] ケーブルリストとネットワーク構成のFromホスト名は一致しました");

            foreach (var key in listCableKeys)
            {
                var listCablePorts = dicFromCableList[key];
                var listDiagramPorts = MyHostNameUsedPorts[key];
                var diffPorts = listCablePorts.Except(listDiagramPorts).Union(listDiagramPorts.Except(listCablePorts));
                if (diffPorts.Count() > 0)
                {
                    isError = true;
                    logger.ZLogError($"ケーブルリストとネットワーク構成にFromホスト名({key})のポート({string.Join(",",diffPorts)})の差分が発見されました");                
                }
                else
                {
                    logger.ZLogTrace($"[checkDiagramVsCableList] ケーブルリストとネットワーク構成のFromホスト名({key})のポートは一致しました");
                }
            }
        }

        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] ネットワーク構成図とケーブルリストの接続ポートの不一致が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] ネットワーク構成図とケーブルリストの接続ポートの一致しました");
        }
        logger.ZLogInformation($"== end ネットワーク構成図とケーブルリストの接続ポートの一致の確認 ==");
    }

    void checkToDeviceAtFromConnect()
    {
        logger.ZLogInformation($"== start 「接続」で対向先の記載の確認 ==");
        bool isError = false;
        string wordConnect = config.Value.WordConnect;
        Dictionary<string,string> dicIgnoreConnectorName = new Dictionary<string, string>();
        string ignoreConnectorNameToAll = config.Value.IgnoreConnectorNameToAll;
        foreach (var ignore in ignoreConnectorNameToAll.Split(','))
        {
            dicIgnoreConnectorName.Add(ignore, "");
        }
        foreach (var device in MyDevicePorts)
        {
            if (device.fromConnect.Equals(wordConnect))
            {
                if (string.IsNullOrEmpty(device.toHostName))
                {
                    if (isNotIgnoreDevice(device.fromConnectorName, dicIgnoreConnectorName))
                    {
                        isError = true;
                        logger.ZLogError($"ケーブルID:{device.fromCableID} From側ホスト名:{device.fromHostName} は({wordConnect})であるが To側ホスト名が記載されていない");
                    }
                    else
                    {
                        logger.ZLogTrace($"[checkToDeviceAtFromConnect] 除外しました ケーブルID:{device.fromCableID} From側コネクター形状:{device.fromConnectorName}");
                    }
                }
                else
                {
                    logger.ZLogTrace($"[checkToDeviceAtFromConnect] ケーブルID:{device.fromCableID} は({wordConnect})であり ({device.toHostName}) と記載あるので問題なし");
                }
            }
            else
            {
                logger.ZLogTrace($"[checkToDeviceAtFromConnect] ケーブルID:{device.fromCableID} は({wordConnect})以外のため確認しない");
            }
        }
        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] 「接続」で対向先の不記載が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] 「接続」で対向先は記載されていました");
        }
        logger.ZLogInformation($"== end 「接続」で対向先の記載の確認 ==");
    }

    void checkConnectXConnect()
    {
        logger.ZLogInformation($"== start 接続される装置間の接続ポートの確認 ==");
        bool isError = false;
        Dictionary<string,string> dicIgnoreDeviceName = new Dictionary<string, string>();
        string ignoreDeviceNameToConnectXConnect = config.Value.IgnoreDeviceNameToConnectXConnect;
        Dictionary<string,string> dicIgnoreConnectorName = new Dictionary<string, string>();
        string ignoreConnectorNameToAll = config.Value.IgnoreConnectorNameToAll;
        foreach (var ignore in ignoreDeviceNameToConnectXConnect.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }
        foreach (var ignore in ignoreConnectorNameToAll.Split(','))
        {
            dicIgnoreConnectorName.Add(ignore, "");
        }

        string wordConnect = config.Value.WordConnect;
        Dictionary<string, string> connectxconnect = new Dictionary<string, string>();
        foreach (var device in MyDevicePorts)
        {
            if (device.fromConnect.Equals(wordConnect))
            {
                if (isNotIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName) && isNotIgnoreDevice(device.fromConnectorName, dicIgnoreConnectorName))
                {
                    try
                    {
                        connectxconnect.Add(device.fromHostName + "&" + device.fromPortName, device.toHostName + "&" + device.toPortName);
                    }
                    catch (System.ArgumentException)
                    {
                        isError = true;
                        logger.ZLogError($"エラー ホスト名とポート番号の組み合わせが重複して記載されています({device.fromHostName + "&" + device.fromPortName}) 2回目の出現ケーブルID:{device.fromCableID}");
                    }
                    catch (System.Exception)
                    {
                        throw;
                    }
                }
                else
                {
                    logger.ZLogTrace($"[checkConnectXConnect] 除外しました ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName} From側コネクター形状:{device.fromConnectorName}");
                }
            }
        }

        foreach (var key in connectxconnect.Keys)
        {
            string toValue = connectxconnect[key];
            if (!connectxconnect.ContainsKey(toValue))
            {
                isError = true;
                logger.ZLogError($"エラー From({key})に対するTo({toValue})が見つかりません");
            }
            else
            {
                string fromValue = connectxconnect[toValue];
                if (!key.Equals(fromValue))
                {
                    isError = true;
                    logger.ZLogTrace($"エラー 元のkey({key})と検索した値から再検索したkey({fromValue})が不一致です");
                }
                else
                {
                    logger.ZLogTrace($"[checkConnectXConnect] チェック通過 From({key})");
                }
            }
        }
        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] 接続される装置間の接続ポートの不一致が発見が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] 接続される装置間の接続ポートの不一致はありませんでした");
        }
        logger.ZLogInformation($"== end 接続される装置間の接続ポートの確認 ==");
    }

    void checkRosette()
    {
        logger.ZLogInformation($"== start ローゼット名の一意の確認 ==");
        bool isError = false;
        Dictionary<string,string> dicDeviceName = new Dictionary<string, string>();
        string deviceNameToRosette = config.Value.DeviceNameToRosette;
        foreach (var device in deviceNameToRosette.Split(','))
        {
            dicDeviceName.Add(device, "");
        }

        string wordConnect = config.Value.WordConnect;
        Dictionary<string,int> dicRosetteName = new Dictionary<string, int>();
        foreach (var device in MyDevicePorts)
        {
            if (device.fromConnect == wordConnect)
            {
                if (isDevice(device.toDeviceName, dicDeviceName))
                {
                    try
                    {
                        dicRosetteName.Add(device.toHostName, device.fromCableID);
                    }
                    catch (System.ArgumentException)
                    {
                        isError = true;
                        logger.ZLogError($"エラー ローゼット名が重複して記載されています({device.toHostName}) 初回の出現ケーブルID:{dicRosetteName[device.toHostName]} 重複回の出現ケーブルID:{device.fromCableID}");
                    }
                    catch (System.Exception)
                    {
                        throw;
                    }
                }
                else
                {
                    logger.ZLogTrace($"[checkRosette] 対象外としました ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName}");
                }
            }
        }

        if (isError)
        {
            isAllPass = false;
            logger.ZLogInformation($"[NG] ローゼット名の重複が発見されました");
        }
        else
        {
            logger.ZLogInformation($"[OK] ローゼット名の重複はありませんでした");
        }
        logger.ZLogInformation($"== end ローゼット名の一意の確認 ==");
    }

    private string getTime()
    {
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, jstTimeZoneInfo).ToString("yyyy-MM-dd'T'HH:mm:sszzz");
    }

    private string getMyFileVersion()
    {
        System.Diagnostics.FileVersionInfo ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
        return ver.InternalName + "(" + ver.FileVersion + ")";
    }

}

//==
public class MyConfig
{
    public string Header {get; set;} = "";

    public string ModelAndPortName {get; set;} = "";
    public string IgnoreModelName {get; set;} = "";
    public int DeviceFromCableIdColumn {get; set;} = -1;
    public int DeviceFromConnectColumn {get; set;} = -1;
    public int DeviceFromFloorNameColumn {get; set;} = -1;
    public int DeviceFromDeviceNameColumn {get; set;} = -1;
    public int DeviceFromDeviceNumberColumn {get; set;} = -1;
    public int DeviceFromModelNameColumn {get; set;} = -1;
    public int DeviceFromHostNameColumn {get; set;} = -1;
    public int DeviceFromPortNameColumn {get; set;} = -1;
    public int DeviceFromConnectorNameColumn {get; set;} = -1;
    public int DeviceToFloorNameColumn {get; set;} = -1;
    public int DeviceToDeviceNameColumn {get; set;} = -1;
    public int DeviceToDeviceNumberColumn {get; set;} = -1;
    public int DeviceToModelNameColumn {get; set;} = -1;
    public int DeviceToHostNameColumn {get; set;} = -1;
    public int DeviceToPortNameColumn {get; set;} = -1;
    public string WordConnect {get; set;} = "";
    public string WordDisconnect {get; set;} = "";
    public string IgnoreDeviceNameToHostNameLength {get; set;} = "";
    public string IgnoreDeviceNameToHostNamePrefix {get; set;} = "";
    public string IgnoreDeviceNameToConnectXConnect {get; set;} = "";
    public string IgnoreConnectorNameToAll {get; set;} = "";
    public string WordDeviceToHostNameList {get; set;} = "";
    public string DeviceNameToRosette {get; set;} = "";
}

public class MyDevicePort
{
    public int fromCableID = -1;
    public string fromConnect = "";

    public string fromFloorName = "";
    public string fromDeviceName = "";
    public string fromDeviceNumber = "";
    public string fromModelName = "";
    public string fromHostName = "";
    public string fromPortName = "";
    public string fromConnectorName = "";

    public string toFloorName = "";
    public string toDeviceName = "";
    public string toDeviceNumber = "";
    public string toModelName = "";
    public string toHostName = "";
    public string toPortName = "";
}