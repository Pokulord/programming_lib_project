#include <iostream>
#include <Windows.h>
#include <bit>
#include <string>
#include <vector>
#include <cmath> // Для функции pow

using namespace std;

#pragma pack(push, 1)

typedef struct {
    BYTE Signature[8];
    GUID CLSID;
    WORD minorVersion;
    WORD majorVersion;
    WORD byteOrder;
    WORD sectorShift;
    WORD miniSectorShift;
    BYTE reserved[6];
    DWORD numOfDirSectors;
    DWORD numOfFATSectors;
    DWORD firstDirSecLoc;
    DWORD TSN;
    DWORD miniStreamCUSize;
    DWORD firstMiniFATSecLoc;
    DWORD numOfMiniFATSectors;
    DWORD firstDIFATSecLoc;
    DWORD numOfDIFATSectors;
    DWORD DIFAT[109];
} CFHeader;

typedef struct {
    WCHAR dirName[32];
    WORD nameLength;
    BYTE objType;
    BYTE colorFlag;
    DWORD leftSibID;
    DWORD rightSibID;
    DWORD childID;
    GUID CLSID;
    DWORD stateBits;
    ULONGLONG creationTime;
    ULONGLONG modifiedTime;
    DWORD startingSecLoc;
    ULONGLONG streamSize;
} DirectoryEntry;

typedef struct {
    WORD vers;
    WORD dt;
    WORD rupBuild;
    WORD rupYear;
    BYTE trash[8];
} BOF;

#pragma pack(pop)

WORD reverse_word(WORD in) {
    WORD b2 = in & 0xFF;
    WORD b1 = (in & 0xFF00) >> 8;
    WORD res = (b2 << 8) | b1;
    return res;
}

DWORD reverse_dword(DWORD in) {
    DWORD b4 = in & 0xFF;
    DWORD b3 = (in & 0xFF00) >> 8;
    DWORD b2 = (in & 0xFF0000) >> 16;
    DWORD b1 = (in & 0xFF000000) >> 24;
    DWORD res = (b4 << 24) | (b3 << 16) | (b2 << 8) | b1;
    return res;
}

bool isEqualArr(BYTE* arr1, BYTE* arr2, DWORD size) {
    for (DWORD i = 0; i < size; i++) {
        if (arr1[i] != arr2[i]) {
            return false;
        }
    }
    return true;
}

int main() {
    setlocale(LC_ALL, "Russian");

    HANDLE fileHandle = CreateFile(L"Книга1.xls", GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (fileHandle == INVALID_HANDLE_VALUE) {
        cout << "Ошибка открытия файла!" << endl;
        return 1;
    }

    DWORD fileSize = GetFileSize(fileHandle, NULL);
    DWORD bytesRead;
    BYTE* buf = new BYTE[fileSize];
    BOOL readOK = ReadFile(fileHandle, buf, fileSize, &bytesRead, NULL);
    if (!readOK) {
        cout << "Ошибка чтения файла!" << endl;
        CloseHandle(fileHandle);
        delete[] buf;
        return -1;
    }

    BYTE trueHeaderSig[] = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
    CFHeader cfh;
    DWORD fileOffset = 0;
    memcpy(&cfh, &buf[fileOffset], sizeof(cfh));
    if (!isEqualArr(trueHeaderSig, cfh.Signature, 8)) {
        wcout << L"Файл не является Compound File!\n";
        delete[] buf;
        CloseHandle(fileHandle);
        return 1;
    }
    fileOffset += sizeof(cfh);

    // Переход к первому DirectoryEntry
    fileOffset += cfh.firstDirSecLoc * static_cast<DWORD>(pow(2, cfh.sectorShift));
    DirectoryEntry de;

    while (fileOffset < fileSize) {
        memcpy(&de, &buf[fileOffset], sizeof(DirectoryEntry));
        if (wstring(de.dirName) == L"Workbook") {
            wcout << L"Workbook stream found!" << endl;
            fileOffset = sizeof(CFHeader) + de.startingSecLoc * static_cast<DWORD>(pow(2, cfh.sectorShift));
            break;
        }
        fileOffset += 128;
    }

    BOF StreamBOF;
    DWORD wbOffset = fileOffset;
    bool foundGlobalSubstream = false;
    vector<DWORD> worksheetOffsets;

    while (fileOffset < fileSize) {
        memcpy(&StreamBOF, &buf[fileOffset], sizeof(BOF));

        // Поиск Globals Substream
        if (!foundGlobalSubstream && StreamBOF.dt == 0x5 && StreamBOF.trash[4] == 6 &&
            (StreamBOF.rupYear == 0x7cc || StreamBOF.rupYear == 0x7cd) && StreamBOF.vers == 0x600) {
            cout << "Globals Substream found! Offset difference to workbook stream BOF: " << (fileOffset - wbOffset) << endl;
            foundGlobalSubstream = true;
            fileOffset += sizeof(BOF);
        }
        // Поиск Worksheet Substream в пределах Globals Substream
        else if (foundGlobalSubstream && StreamBOF.dt == 0x10 && StreamBOF.vers == 0x600) {
            DWORD wsOffset = fileOffset;
            cout << "Worksheet Substream found at offset: " << wsOffset << endl;
            worksheetOffsets.push_back(wsOffset);
            fileOffset += sizeof(BOF);
        }
        else {
            fileOffset += 1;
        }
    }

    // Вывод всех найденных смещений для Worksheet Substream
    if (!worksheetOffsets.empty()) {
        cout << "Всего найдено Worksheet Substream: " << worksheetOffsets.size() << endl;
        for (size_t i = 0; i < worksheetOffsets.size(); i++) {
            cout << "Worksheet Substream #" << (i + 1) << " offset: " << worksheetOffsets[i] << endl;
        }
    }
    else {
        cout << "Worksheet Substream не найдены!" << endl;
    }

    CloseHandle(fileHandle);
    delete[] buf;
    return 0;
}
