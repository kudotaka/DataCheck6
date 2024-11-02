using System.Collections.Generic;
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
        if (!File.Exists(excelpath))
        {
            logger.ZLogError($"target excel file is missing.");
            return;
        }
        logger.ZLogInformation($"[command] arg excel:{excelpath}, prefix:{prefix}");

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

        foreach (var sheet in sheets)
        {
            int lastUsedRowNumber = sheet.LastRowUsed() == null ? 0 : sheet.LastRowUsed().RowNumber();
            int lastUsedColumNumber = sheet.LastColumnUsed() == null ? 0 : sheet.LastColumnUsed().ColumnNumber();
            logger.ZLogInformation($"name:{sheet.Name}, lastUsedRow:{lastUsedRowNumber}, lastUsedColum:{lastUsedColumNumber}");
//            logger.ZLogInformation($"name:{sheet.Name}, deviceToHostNameColumn:{deviceToHostNameColumn}, deviceToPortNameColumn:{deviceToPortNameColumn}");

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

//== print
//        printMyDevicePorts();

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
    }

    private void printMyDevicePorts()
    {
        foreach (var device in MyDevicePorts)
        {
            logger.ZLogDebug($"CableID:{device.fromCableID}, connect:{device.fromConnect}, fromHost:{device.fromHostName} fromPort:{device.fromPortName} toHost:{device.toHostName} toPort:{device.toPortName}");
        }
    }

    private void checkDuplicateCableId()
    {
        Dictionary<int, string> cableId = new Dictionary<int, string>();
        foreach (var device in MyDevicePorts)
        {
            try
            {
                cableId.Add(device.fromCableID, device.fromHostName + device.fromPortName);
            }
            catch (System.ArgumentException)
            {
                logger.ZLogError($"[checkDuplicateCableId] CableId is duplicate! at cableId:{device.fromCableID}");
            }
            catch (System.Exception)
            {
                throw;
            }
        }
    }

    private bool isIgnoreDevice(string device, Dictionary<string,string> dicIgnore)
    {
        return !dicIgnore.ContainsKey(device);
    }

    private void checkHostNameLength()
    {
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
                if (isIgnoreDevice(device.fromDeviceName, dicIgnoreDeviceName))
                {
                    logger.ZLogError($"[checkHostNameLength] hostname(from) length is miss! at cableId:{device.fromCableID} fromDeivcename:{device.fromDeviceName}");
                }
            }

            if (device.fromConnect == wordConnect)
            {
                if (device.toHostName.Length != deviceHostNameLength)
                {
                    if (device.toHostName.Length != rosetteHostNameLength)
                    {
                        if (isIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName))
                        {
                            logger.ZLogError($"[checkHostNameLength] hostname(to) length is miss! at cableId:{device.fromCableID} toDevicename:{device.toDeviceName}");
                        }
                    }
                }
            }
        }
    }

    private void checkHostNamePrefix(string prefix)
    {
        Dictionary<string,string> dicIgnoreDeviceName = new Dictionary<string, string>();
        string ignoreDeviceNameToHostNamePrefix = config.Value.IgnoreDeviceNameToHostNamePrefix;
        foreach (var ignore in ignoreDeviceNameToHostNamePrefix.Split(','))
        {
            dicIgnoreDeviceName.Add(ignore, "");
        }

        string wordConnect = config.Value.WordConnect;
        foreach (var device in MyDevicePorts)
        {
            //hostname.StartsWith(prefix)
            if (!device.fromHostName.StartsWith(prefix))
            {
                if (isIgnoreDevice(device.fromDeviceName, dicIgnoreDeviceName))
                {
                    logger.ZLogError($"[checkHostNamePrefix] hostname(from) prefix is not match! at cableId:{device.fromCableID} prefix:{prefix} fromHostname:{device.fromHostName}");
                }
            }

            if (device.fromConnect == wordConnect)
            {
                if (!device.toHostName.StartsWith(prefix))
                {
                    if (isIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName))
                    {
                        logger.ZLogError($"[checkHostNamePrefix] hostname(to) length is not match! at cableId:{device.fromCableID} prefix:{prefix} toHostname:{device.toHostName}");
                    }
                }
            }
        }
    }

    void checkToDeviceAtFromConnect()
    {
        string wordConnect = config.Value.WordConnect;
        foreach (var device in MyDevicePorts)
        {
            if (!string.IsNullOrEmpty(device.toHostName))
            {
                if (!device.fromConnect.Equals(wordConnect))
                {
                    logger.ZLogError($"[checkToDeviceAtFromConnect] device(to) is alive.but not Connect! at cableId:{device.fromCableID} fromHostname:{device.fromHostName}");
                }
            }
        }
    }

    void checkConnectXConnect()
    {
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
            if (device.fromConnect.Equals(wordConnect) && isIgnoreDevice(device.toDeviceName, dicIgnoreDeviceName))
            {
                try
                {
                    connectxconnect.Add(device.fromHostName + "&" + device.fromPortName, device.toHostName + "&" + device.toPortName);
                }
                catch (System.ArgumentException)
                {
                    logger.ZLogError($"[checkConnectXConnect] duplicate! at cableId:{device.fromCableID}");
                }
                catch (System.Exception)
                {
                    throw;
                }
            }
        }

        foreach (var key in connectxconnect.Keys)
        {
//            logger.ZLogInformation($"{key}");
            string toValue = connectxconnect[key];
            if (!connectxconnect.ContainsKey(toValue))
            {
                logger.ZLogError($"[checkConnectXConnect] toKey:{toValue} is miss! fromKey:{key}");
            }
        }
    }

    private string getTime()
    {
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, jstTimeZoneInfo).ToString("yyyy-MM-dd'T'HH:mm:sszzz");
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