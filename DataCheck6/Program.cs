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

    readonly ILogger<DataCheckApp> logger;
    readonly IOptions<MyConfig> config;

    List<MyDevicePort> MyDevicePorts = new List<MyDevicePort>();

    public DataCheckApp(ILogger<DataCheckApp> logger,IOptions<MyConfig> config)
    {
        this.logger = logger;
        this.config = config;
    }

//    [Command("")]
    public void Check(string excelpath, string prefix)
    {
//== start
        logger.ZLogInformation($"==== tool {getMyFileVersion()} ====");
        
        if (!File.Exists(excelpath))
        {
            logger.ZLogError($"target excel file is missing.");
            return;
        }

        FileStream fs = new FileStream(excelpath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbook = new XLWorkbook(fs);
        IXLWorksheets sheets = xlWorkbook.Worksheets;

//== init
        int deviceFromCableIdColumn = config.Value.DeviceFromCableIdColumn;
        int deviceFromConnectColumn = config.Value.DeviceFromConnectColumn;
        int deviceFromDeviceNameColumn = config.Value.DeviceFromDeviceNameColumn;
        int deviceFromHostNameColumn = config.Value.DeviceFromHostNameColumn;
        int deviceFromModelNameColumn = config.Value.DeviceFromModelNameColumn;
        int deviceFromPortNameColumn = config.Value.DeviceFromPortNameColumn;
        int deviceToDeviceNameColumn = config.Value.DeviceToDeviceNameColumn;
        int deviceToModelNameColumn = config.Value.DeviceToModelNameColumn;
        int deviceToHostNameColumn = config.Value.DeviceToHostNameColumn;
        int deviceToPortNameColumn = config.Value.DeviceToPortNameColumn;
        string wordConnect = config.Value.WordConnect;
        string wordDisconnect = config.Value.WordDisconnect;
        int hostNameLength = config.Value.DeviceHostNameLength;
        string ignoreDeviceNameToHostNameLength = config.Value.IgnoreDeviceNameToHostNameLength;
        int rosetteHostNameLength = config.Value.RosetteHostNameLength;
        string ignoreDeviceNameToHostNamePrefix = config.Value.IgnoreDeviceNameToHostNamePrefix;
        string ignoreDeviceNameToConnectXConnect = config.Value.IgnoreDeviceNameToConnectXConnect;

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
                            tmpDevicePort.fromDeviceName = sheet.Cell(r, deviceFromDeviceNameColumn).Value.ToString();
                            tmpDevicePort.fromHostName = sheet.Cell(r, deviceFromHostNameColumn).Value.ToString();
                            tmpDevicePort.fromModelName = sheet.Cell(r, deviceFromModelNameColumn).Value.ToString();
                            tmpDevicePort.fromPortName = sheet.Cell(r, deviceFromPortNameColumn).Value.ToString();
                            tmpDevicePort.toDeviceName = sheet.Cell(r, deviceToDeviceNameColumn).Value.ToString();
                            tmpDevicePort.toModelName = sheet.Cell(r, deviceToModelNameColumn).Value.ToString();
                            tmpDevicePort.toHostName = sheet.Cell(r, deviceToHostNameColumn).Value.ToString();
                            tmpDevicePort.toPortName = sheet.Cell(r, deviceToPortNameColumn).Value.ToString();
                            MyDevicePorts.Add(tmpDevicePort);
                        }
                    }
                }
            }
        }

//== print
        printMyDevicePorts();

//== check duplicate CableID
        checkDuplicateCableId();

//== check toDevice --> Connect
        checkToDeviceAtFromConnect();

//== check hostname count
        checkHostNameLength();

//== check hostname prefix
        checkHostNamePrefix(prefix);

//== check 
        checkConnectXConnect();

//== finish
        if (isAllPass)
        {
            logger.ZLogInformation($"== [Congratulations!] すべての確認項目をパスしました ==");
        }
        logger.ZLogInformation($"==== tool finish ====");

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

    private void checkHostNameLength()
    {
        logger.ZLogInformation($"== start ホスト名の長さの確認 ==");
        bool isError = false;
        Dictionary<string,string> dicIgnoreDeviceName = new Dictionary<string, string>();
        string ignoreDeviceNameToHostNameLength = config.Value.IgnoreDeviceNameToHostNameLength;
        foreach (var ignore in ignoreDeviceNameToHostNameLength.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }

        string wordConnect = config.Value.WordConnect;
        int deviceHostNameLength = config.Value.DeviceHostNameLength;
        int rosetteHostNameLength = config.Value.RosetteHostNameLength;
        foreach (var device in MyDevicePorts)
        {
            if (device.fromHostName.Length != deviceHostNameLength)
            {
                if (isNotIgnoreDevice(device.fromDeviceName, dicIgnoreDeviceName))
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} From側デバイス名:{device.fromDeviceName} From側ホスト名:{device.fromHostName}");
                }
                else
                {
                    logger.ZLogTrace($"[checkHostNameLength] 除外しました ケーブルID:{device.fromCableID} From側デバイス名:{device.fromDeviceName}");
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
                        if (isNotIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName))
                        {
                            isError = true;
                            logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName} To側ホスト名:{device.toHostName}");
                        }
                        else
                        {
                            logger.ZLogTrace($"[checkHostNameLength] 除外しました ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName}");
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
        foreach (var ignore in ignoreDeviceNameToHostNamePrefix.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }

        string wordConnect = config.Value.WordConnect;
        foreach (var device in MyDevicePorts)
        {
            if (!device.fromHostName.StartsWith(prefix))
            {
                if (isNotIgnoreDevice(device.fromDeviceName, dicIgnoreDeviceName))
                {
                    isError = true;
                    logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} 接頭語:{prefix} From側ホスト名:{device.fromHostName}");
                }
                else
                {
                    logger.ZLogTrace($"[checkHostNamePrefix] 除外しました ケーブルID:{device.fromCableID} From側デバイス名:{device.fromDeviceName}");
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
                    if (isNotIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName))
                    {
                        isError = true;
                        logger.ZLogError($"不一致エラー ケーブルID:{device.fromCableID} 接頭語:{prefix} To側ホスト名:{device.toHostName}");
                    }
                    else
                    {
                        logger.ZLogTrace($"[checkHostNamePrefix] 除外しました ケーブルID:{device.fromCableID} To側デバイス名:{device.toDeviceName}");
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

    void checkToDeviceAtFromConnect()
    {
        logger.ZLogInformation($"== start 「接続」で対向先の記載の確認 ==");
        bool isError = false;
        string wordConnect = config.Value.WordConnect;
        foreach (var device in MyDevicePorts)
        {
            if (device.fromConnect.Equals(wordConnect))
            {
                if (string.IsNullOrEmpty(device.toHostName))
                {
                    isError = true;
                    logger.ZLogError($"ケーブルID:{device.fromCableID} From側ホスト名:{device.fromHostName} は({wordConnect})であるが To側ホスト名が記載されていない");
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
        foreach (var ignore in ignoreDeviceNameToConnectXConnect.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }

        string wordConnect = config.Value.WordConnect;
        Dictionary<string, string> connectxconnect = new Dictionary<string, string>();
        foreach (var device in MyDevicePorts)
        {
            if (device.fromConnect.Equals(wordConnect))
            {
                if (isNotIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName))
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
                    logger.ZLogTrace($"[checkConnectXConnect] 除外しました cableId:{device.fromCableID} toDevicename:{device.toDeviceName}");
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

    public int DeviceFromCableIdColumn {get; set;} = -1;
    public int DeviceFromConnectColumn {get; set;} = -1;
    public int DeviceFromDeviceNameColumn {get; set;} = -1;
    public int DeviceFromModelNameColumn {get; set;} = -1;
    public int DeviceFromHostNameColumn {get; set;} = -1;
    public int DeviceFromPortNameColumn {get; set;} = -1;
    public int DeviceToDeviceNameColumn {get; set;} = -1;
    public int DeviceToModelNameColumn {get; set;} = -1;
    public int DeviceToHostNameColumn {get; set;} = -1;
    public int DeviceToPortNameColumn {get; set;} = -1;
    public string WordConnect {get; set;} = "";
    public string WordDisconnect {get; set;} = "";
    public int DeviceHostNameLength {get; set;} = 0;
    public string IgnoreDeviceNameToHostNameLength {get; set;} = "";
    public int RosetteHostNameLength {get; set;} = 0;
    public string IgnoreDeviceNameToHostNamePrefix {get; set;} = "";
    public string IgnoreDeviceNameToConnectXConnect {get; set;} = "";
}

public class MyDevicePort
{
    public int fromCableID = -1;
    public string fromConnect = "";

    public string fromDeviceName = "";
    public string fromModelName = "";
    public string fromHostName = "";
    public string fromPortName = "";

    public string toDeviceName = "";
    public string toModelName = "";
    public string toHostName = "";
    public string toPortName = "";
}