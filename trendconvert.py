import sys
import struct
import math
from datetime import datetime, timedelta
import argparse
from openpyxl import Workbook, workbook
import csv


class MasterHeader:
    Title = None
    ID = None
    Type = None
    Version = None
    Max_nr_files = None
    Files_created = None
    Next = None
    Addon = None
    Datafile_names = None
    Data_headers = None


class Header:
    ID = None
    Type = None
    Version = None
    StartEvNo = None
    LogName = None
    Mode = None
    Area = None
    Priv = None
    FileType = None
    SamplePediod = None
    EngUnits = None
    Format = None
    StartTime = None
    EndTime = None
    DataLength = None
    FilePointer = None
    EndEvNo = None


class EngScale:
    RawZero = None
    RawFull = None
    EngZero = None
    EngFull = None


def parseArgs():
    parser = argparse.ArgumentParser("trendconvert")
    parser.add_argument("file", type=str, help="Filename (TRENDFILE.HST)")
    parser.add_argument(
        "-s",
        default=False,
        action="store_true",
        help="Strip directories from HST data filenames. (If files are moved from original directory.)",
    )
    parser.add_argument(
        "-o",
        metavar="TYPE",
        type=str,
        choices=["xls", "csv"],
        default="xls",
        help="Output file type (xls,csv)",
    )
    parser.add_argument(
        "-e",
        default=False,
        action="store_true",
        help="Examine master header for dates in data files.",
    )
    parser.add_argument(
        "-start", type=str, metavar="DATE", help="Start date (YYYY-MM-DD)"
    )
    parser.add_argument("-stop", type=str, metavar="DATE", help="End date (YYYY-MM-DD)")
    parser.add_argument("-f", type=int, metavar="NUM", help="Select file to export.")
    parser.add_argument(
        "-d",
        default=True,
        action="store_false",
        help="Do not discard invalid values from samples.",
    )
    parser.add_argument(
        "-p",
        type=int,
        default=1,
        metavar="NUM",
        help="Number of decimals shown in values. (Default: 1)",
    )
    args = parser.parse_args()
    return args


def readMasterHeader(f):
    m = MasterHeader()
    m.Title = f.read(128).decode("cp1252").rstrip("\x00")
    m.ID = f.read(8).decode("cp1252").rstrip("\x00")
    m.Type = int.from_bytes(f.read(2), "little", signed=False)  # UINT default=False
    m.Version = int.from_bytes(f.read(2), "little")  # UINT 5=2, byte 6=8 byte
    tAlign = f.read(4)  # Alignment
    tMode = f.read(4)  # Uint set to 0
    m.Max_nr_files = int.from_bytes(f.read(2), "little")
    m.Files_created = int.from_bytes(f.read(2), "little")
    m.Next = f.read(2)
    m.Addon = f.read(2)
    tAlign = f.read(20)
    m.Datafile_names = []
    m.Data_headers = []
    return m


def readOldTypeHeaders(m, f):
    for x in range(m.Files_created):
        filename = f.read(144).decode("cp1252")
        m.Datafile_names.append(filename.rstrip("\x00"))
        h = Header()
        h.ID = f.read(8).decode("cp1252").rstrip("\x00")
        h.Type = int.from_bytes(f.read(2), "little")
        h.Version = int.from_bytes(f.read(2), "little")
        h.StartEvNo = int.from_bytes(f.read(4), "little", signed=True)
        # tAlign = f.read()
        h.LogName = f.read(80).decode("cp1252").rstrip("\x00")
        h.Mode = int.from_bytes(f.read(4), "little")
        h.Area = int.from_bytes(f.read(2), "little")
        h.Priv = int.from_bytes(f.read(2), "little")
        h.FileType = int.from_bytes(f.read(2), "little")
        h.SamplePediod = int.from_bytes(f.read(4), "little")
        h.EngUnits = f.read(8).decode("cp1252").rstrip("\x00")
        h.Format = int.from_bytes(f.read(4), "little")
        h.StartTime = datetime.fromtimestamp(int.from_bytes(f.read(4), "little"))
        h.EndTime = datetime.fromtimestamp(int.from_bytes(f.read(4), "little"))
        h.DataLength = int.from_bytes(f.read(4), "little")
        h.FilePointer = int.from_bytes(f.read(4), "little")
        h.EndEvNo = int.from_bytes(f.read(4), "little", signed=True)
        tAlign = f.read(2)
        m.Data_headers.append(h)


def readNewTypeHeaders(m, f):
    for x in range(m.Files_created):
        filename = f.read(272).decode("cp1252")
        m.Datafile_names.append(filename.rstrip("\x00"))
        h = Header()
        h.ID = f.read(8).decode("cp1252").rstrip("\x00")
        h.Type = int.from_bytes(f.read(2), "little")
        h.Version = int.from_bytes(f.read(2), "little")
        h.StartEvNo = int.from_bytes(f.read(8), "little", signed=True)
        tAlign = f.read(12)
        h.LogName = f.read(80).decode("cp1252").rstrip("\x00")
        h.Mode = int.from_bytes(f.read(4), "little")
        h.Area = int.from_bytes(f.read(2), "little")
        h.Priv = int.from_bytes(f.read(2), "little")
        h.FileType = int.from_bytes(f.read(2), "little")
        h.SamplePediod = int.from_bytes(f.read(4), "little")
        h.EngUnits = f.read(8).decode("cp1252").rstrip("\x00")
        h.Format = int.from_bytes(f.read(4), "little")
        h.StartTime = datetime(1601, 1, 1) + timedelta(
            microseconds=int.from_bytes(f.read(8), "little") / 10
        )
        h.EndTime = datetime(1601, 1, 1) + timedelta(
            microseconds=int.from_bytes(f.read(8), "little") / 10
        )
        h.DataLength = int.from_bytes(f.read(4), "little")
        h.FilePointer = int.from_bytes(f.read(4), "little")
        h.EndEvNo = int.from_bytes(f.read(8), "little", signed=True)
        tAlign = f.read(6)
        m.Data_headers.append(h)


def readOldDataHeader(f):
    h = Header()
    h.ID = f.read(8).decode("cp1252").rstrip("\x00")
    h.Type = int.from_bytes(f.read(2), "little")
    h.Version = int.from_bytes(f.read(2), "little")
    h.StartEvNo = int.from_bytes(f.read(4), "little", signed=True)
    h.LogName = f.read(80).decode("cp1252").rstrip("\x00")
    h.Mode = int.from_bytes(f.read(4), "little")
    h.Area = int.from_bytes(f.read(2), "little")
    h.Priv = int.from_bytes(f.read(2), "little")
    h.FileType = int.from_bytes(f.read(2), "little")
    h.SamplePediod = int.from_bytes(f.read(4), "little")
    h.EngUnits = f.read(8).decode("cp1252").rstrip("\x00")
    h.Format = int.from_bytes(f.read(4), "little")
    h.StartTime = datetime.fromtimestamp(int.from_bytes(f.read(4), "little"))
    h.EndTime = datetime.fromtimestamp(int.from_bytes(f.read(4), "little"))
    h.DataLength = int.from_bytes(f.read(4), "little")
    h.FilePointer = int.from_bytes(f.read(4), "little")
    h.EndEvNo = int.from_bytes(f.read(4), "little", signed=True)
    tAlign = f.read(2)
    return h


def readNewDataHeader(f):
    h = Header()
    h.ID = f.read(8).decode("cp1252").rstrip("\x00")
    h.Type = int.from_bytes(f.read(2), "little")
    h.Version = int.from_bytes(f.read(2), "little")
    h.StartEvNo = int.from_bytes(f.read(8), "little", signed=True)
    tAlign = f.read(12)
    h.LogName = f.read(80).decode("cp1252").rstrip("\x00")
    h.Mode = int.from_bytes(f.read(4), "little")
    h.Area = int.from_bytes(f.read(2), "little")
    h.Priv = int.from_bytes(f.read(2), "little")
    h.FileType = int.from_bytes(f.read(2), "little")
    h.SamplePediod = int.from_bytes(f.read(4), "little")
    h.EngUnits = f.read(8).decode("cp1252").rstrip("\x00")
    h.Format = int.from_bytes(f.read(4), "little")
    h.StartTime = datetime(1601, 1, 1) + timedelta(
        microseconds=int.from_bytes(f.read(8), "little") / 10
    )
    h.EndTime = datetime(1601, 1, 1) + timedelta(
        microseconds=int.from_bytes(f.read(8), "little") / 10
    )
    h.DataLength = int.from_bytes(f.read(4), "little")
    h.FilePointer = int.from_bytes(f.read(4), "little")
    h.EndEvNo = int.from_bytes(f.read(8), "little", signed=True)
    tAlign = f.read(6)
    return h


def examineDataFiles(m: MasterHeader):
    x = 0
    if m.Version == 5:
        type = "Type: 2 byte"
    else:
        type = "Type: 8 byte"
    print(
        type,
        "| Maximum number of files:",
        m.Max_nr_files,
        "| Files created:",
        m.Files_created,
    )
    for d in m.Data_headers:
        if m.Version == 5:
            print(
                "File:",
                x,
                m.Datafile_names[x],
                "| Start:",
                d.StartTime,
                "| End:",
                d.EndTime,
                "| Samples:",
                d.DataLength,
            )
            x += 1
        else:
            print(
                "File:",
                x,
                m.Datafile_names[x],
                "Start:",
                d.StartTime,
                "End:",
                d.EndTime,
                "| Samples:",
                d.DataLength,
            )
            x += 1
    sys.exit(0)


def readScales(f):
    e = EngScale()
    e.RawZero = struct.unpack("f", f.read(4))[0]
    e.RawFull = struct.unpack("f", f.read(4))[0]
    e.EngZero = struct.unpack("f", f.read(4))[0]
    e.EngFull = struct.unpack("f", f.read(4))[0]
    return e


def calcValue(e: EngScale, meas, precision):
    v = e.EngZero + ((meas - 0) / (32000)) * (e.EngFull - e.EngZero)
    return round(v, precision)


def stripDirectories(m: MasterHeader):
    x = 0
    for d in m.Datafile_names:
        path = d.split("\\")
        file = path[-1]
        m.Datafile_names[x] = file
        x += 1


def selectDataFiles(m: MasterHeader, startTime, stopTime):
    x = 0
    data = []
    for d in m.Data_headers:
        if d.StartTime <= startTime and d.EndTime > startTime:
            data.append(x)
        elif d.StartTime < stopTime and d.EndTime > stopTime:
            data.append(x)
        x += 1
    # print(data)
    return data


def main():
    args = parseArgs()
    if args.file.split(".")[-1].upper() != "HST":
        print("Check filename. Use .HST")
        sys.exit(2)

    # Init lists for master header
    with open(args.file, "rb") as f:
        # read MASTERHEADER
        m = readMasterHeader(f)
        # Start parsing HSTFILEHEADER
        if m.Version == 5:
            readOldTypeHeaders(m, f)
        if m.Version == 6:
            readNewTypeHeaders(m, f)
    f.close()

    if args.s:
        stripDirectories(m)

    # Show summary of files if desired by user
    if args.e:
        examineDataFiles(m)

    # If -start or -stop set, make time objects
    try:
        if args.start:
            startTime = datetime.strptime(args.start, "%Y-%m-%d")
        if args.stop:
            stopTime = datetime.strptime(args.stop, "%Y-%m-%d")
    except Exception as e:
        print("Invalid date format. Use YYYY-MM-DD")
        sys.exit(0)

    # Select data file to read
    if args.start and args.stop:
        datalist = selectDataFiles(m, startTime, stopTime)

    # if specific file is defined
    if args.f:
        datalist = []
        datalist.append(args.f)
    if not datalist:
        datalist = range(0, m.Files_created - 1)
    # Open data files
    for data in datalist:
        if args.o.lower() == "xls":
            wb = Workbook()
            ws = wb.active
        if args.o.lower() == "csv":
            file = open(
                m.Datafile_names[data].split("\\")[-1].replace(".", "_") + ".csv",
                "w",
                newline="",
            )
            writer = csv.writer(file)
            writer.writerow(["Time", "Value"])
        with open(m.Datafile_names[data], "rb") as f:
            dTitle = f.read(112).decode("cp1252")
            e = readScales(f)

            if m.Data_headers[data].Version == 5:
                h = readOldDataHeader(f)
            if m.Data_headers[data].Version == 6:
                h = readNewDataHeader(f)

            # Reading samples
            if m.Data_headers[data].Version == 5:
                sp = h.SamplePediod / 1000
                # for x in range(20):
                x = 0
                while 1:
                    bytes = f.read(2)
                    if not bytes:
                        break
                    value = int.from_bytes(bytes, "little", signed=True)
                    if value == -32001 or value == -32002:
                        x += 1
                        continue
                    realtime = h.StartTime + timedelta(seconds=sp * x)
                    if args.o.lower() == "xls":
                        if args.start and args.stop:
                            if realtime >= startTime and realtime <= stopTime:
                                ws.append([realtime, calcValue(e, value, args.p)])
                        else:
                            ws.append([realtime, calcValue(e, value, args.p)])
                    if args.o.lower() == "csv":
                        if args.start and args.stop:
                            if realtime >= startTime and realtime <= stopTime:
                                writer.writerow([realtime, calcValue(e, value, args.p)])
                        else:
                            writer.writerow([realtime, calcValue(e, value, args.p)])
                    x += 1
            # 8 byte samples
            else:
                x = 0
                while 1:
                    bytes = f.read(8)
                    if not bytes:
                        break
                    value = struct.unpack("@d", bytes)[0]
                    invalid = int.from_bytes(bytes, "little", signed=True)
                    if args.d:
                        if invalid == 4294949819 or invalid == 4294945450:
                            x += 1
                            continue
                    realtime = h.StartTime + timedelta(
                        microseconds=h.SamplePediod * 1000 * x
                    )
                    if args.d and math.isnan(value):
                        x += 1
                        continue
                    else:
                        if args.o.lower() == "xls":
                            if args.start and args.stop:
                                if realtime >= startTime and realtime <= stopTime:
                                    ws.append([realtime, round(value, args.p)])
                            else:
                                ws.append([realtime, round(value, args.p)])
                        if args.o.lower() == "csv":
                            if args.start and args.stop:
                                if realtime >= startTime and realtime <= stopTime:
                                    writer.writerow([realtime, round(value, args.p)])
                            else:
                                writer.writerow([realtime, round(value, args.p)])
                        x += 1
        if args.o.lower() == "xls":
            wb.save(m.Datafile_names[data].split("\\")[-1].replace(".", "_") + ".xlsx")
        if args.o.lower() == "csv":
            file.close()
        f.close()


if __name__ == "__main__":
    main()
