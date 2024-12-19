#pragma once
#ifndef xls_funcs
#define xls_funcs
#include <cmath>
#include <cstring>
#include <fstream>
#include <iostream>
#include <string>
#include <vector>


using namespace std;
#ifdef _WIN32
typedef unsigned short WORD;
typedef unsigned long DWORD;
typedef unsigned long long ULONGLONG;
#endif // _WIN32
#ifdef __linux__
typedef unsigned short WORD;
typedef unsigned int DWORD;
typedef unsigned long ULONGLONG;
#endif // __linux__

#pragma pack(push, 1)

struct CFHeader {
	unsigned char Siganture[8];
	unsigned char CLSID[16];
	WORD minorVersion;
	WORD majorVersion;
	WORD charOrder;
	WORD sectorShift;
	WORD miniSectorShift;
	unsigned char reserved[6];
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
};

struct DirectoryEntry {
	char16_t dirName[32];
	WORD nameLength;
	unsigned char objType;
	unsigned char colorFlag;
	DWORD leftSibID;
	DWORD rightSibID;
	DWORD childID;
	unsigned char CLSID[16];
	DWORD stateBits;
	ULONGLONG creationTime;
	ULONGLONG modifiedTime;
	DWORD startingSecLoc;
	ULONGLONG streamSize;
};

struct RecordHead {
	WORD type;
	WORD size;
};

struct WsBool {
	unsigned char fShowAutoBreaks : 1;
	unsigned char reserved1 : 3;
	unsigned char fDialog : 1;
	unsigned char fApplyStyles : 1;
	unsigned char fRowSumsBelow : 1;
	unsigned char fColSumsRight : 1;
	unsigned char fFitToPage : 1;
	unsigned char reserved2 : 1;
	unsigned char unused : 2;
	unsigned char fSyncHoriz : 1;
	unsigned char fSyncVert : 1;
	unsigned char fAltExprEval : 1;
	unsigned char fAltFormulaEntry : 1;
};

struct SXLUS {
	unsigned char cch;
	unsigned char fhighByte : 1;
	unsigned char reserved : 7;
	vector<char> rgb;
};

struct BoundSheet8 {
	DWORD pos;
	unsigned char trash;
	unsigned char dt;
	SXLUS name;
};

struct XLURES {
	WORD cch;
	unsigned char fhighByte : 1;
	unsigned char reserved1 : 1;
	unsigned char fExtSt : 1;
	unsigned char fRichSt : 1;
	unsigned char reserved2 : 4;
	WORD cRun = 0;
	DWORD cbExtRst = 0;
	vector<char> rgb;
};

struct SST {
	DWORD cstTotal;
	DWORD cstUnique;
	vector<XLURES> strings;
};

struct Index {
	DWORD reserved;
	DWORD rwMic;
	DWORD rwMac;
	DWORD ibXF;
	vector<DWORD> rgibRw;
};

struct DBCell {
	DWORD dbRtrw;
	vector<WORD> rgdb;
};

struct Cell {
	WORD rw;
	WORD col;
	WORD ixfe;
};

struct LabelSst {
	Cell cell;
	DWORD isst;
};

#pragma pack(pop)

wstring getCellIndex(int columnIndex, int rowIndex) {
	std::wstring columnLabel;
	while (columnIndex > 0) {
		int remainder = (columnIndex - 1) % 26;
		columnLabel = static_cast<wchar_t>(L'A' + remainder) + columnLabel;
		columnIndex = (columnIndex - 1) / 26;
	}
	return columnLabel + std::to_wstring(rowIndex);
}

bool isEqualArr(unsigned char* arr1, unsigned char* arr2, DWORD size) {
	for (int i = 0; i < size; i++) {
		if (arr1[i] != arr2[i]) {
			return false;
		}
	}
	return true;
}

inline void ReadSector(void* dst, char* buf, DWORD secNum, DWORD size) {
	memcpy(dst, &buf[(secNum + 1) * size], size);
}

void getDIFATChain(DWORD* dst, CFHeader* cfh, char* buf, DWORD secSize) {
	dst[0] = cfh->firstDIFATSecLoc;
	for (int i = 1; i < cfh->numOfDIFATSectors; i++) {
		DWORD next;
		memcpy(&next, &buf[dst[i - 1] * (secSize + 1) + secSize - 4], 4);
		if (next == 0xfffffffe) {
			break;
		}
		dst[i] = next;
	}
}

void getFATChain(DWORD* dst, DWORD* difCh, CFHeader* cfh, char* buf, DWORD secSize) {
	for (int i = 0; i < 109; i++) {
		if (cfh->DIFAT[i] == 0xffffffff) {
			break;
		}
		dst[i] = cfh->DIFAT[i];
	}
	if (cfh->numOfFATSectors > 109) {
		for (int i = 0; i < cfh->numOfDIFATSectors; i++) {
			DWORD* difat = new DWORD[secSize / 4];
			ReadSector(difat, buf, difCh[i], secSize);
			for (int j = 0; j < (secSize / 4) - 1; j++) {
				if (difat[j] == 0xffffffff) {
					break;
				}
				dst[109 + (i * secSize) + j] = difat[j];
			}
			delete[] difat;
		}
	}
}

void getWorkbookChain(DWORD* dst, DWORD* fatCh, DirectoryEntry* de, DWORD chainSize, DWORD secSize, char* buf) {
	dst[0] = de->startingSecLoc;
	for (int i = 1; i < chainSize; i++) {
		DWORD curFATSecIndex = dst[i - 1] / (secSize / 4);
		DWORD* fat = new DWORD[secSize / 4];
		ReadSector(fat, buf, fatCh[curFATSecIndex], secSize);
		dst[i] = fat[dst[i - 1] % (secSize / 4)];
		delete[] fat;
	}
}

ifstream openFile(const char* filename) {
	ifstream ifs(filename, ios::binary);
	if (!ifs.is_open()) {
		wcout << L"������ �������� �����!" << endl;
		exit(1);
	}
	return ifs;
}

char* getData(ifstream* file) {
	file->seekg(0, ios::end);
	DWORD fileSize = file->tellg();
	file->seekg(0, ios::beg);
	char* buf = new char[fileSize];
	file->read(buf, fileSize);
	file->close();
	return buf;
}

CFHeader* readHeader(char* buf) {
	CFHeader* cfh = new CFHeader;
	memcpy(cfh, &buf[0], 512);
	return cfh;
}

void checkSig(CFHeader* cfh) {
	unsigned char trueHeaderSig[] = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
	if (!isEqualArr(trueHeaderSig, cfh->Siganture, 8)) {
		wcout << L"������������ ��������� �����!" << endl;;
		exit(1);
	}
	else {
		wcout << L"��������� ����� �����!" << endl;
	}
}

DWORD countDESectors(DWORD fstDELoc, char* buf, DWORD* fatCh, DWORD secSize) {
	DWORD k = 1;
	DWORD i = fstDELoc;
	while (1) {
		DWORD curFATSecIndex = i / (secSize / 4);
		DWORD* fat = new DWORD[secSize / 4];
		ReadSector(fat, buf, fatCh[curFATSecIndex], secSize);
		if (fat[i % (secSize / 4)] == 0xfffffffe) {
			break;
		}
		i = fat[i % (secSize / 4)];
		k++;
		delete[] fat;
	}
	return k;
}

DWORD* getDEChain(DWORD fstDELoc, char* buf, DWORD* fatCh, DWORD secSize, DWORD k) {
	DWORD* deChain = new DWORD[k];
	deChain[0] = fstDELoc;
	for (int i = 1; i < k; i++) {
		DWORD curFATSecIndex = deChain[i - 1] / (secSize / 4);
		DWORD* fat = new DWORD[secSize / 4];
		ReadSector(fat, buf, fatCh[curFATSecIndex], secSize);
		deChain[i] = fat[deChain[i - 1] % (secSize / 4)];
		delete[] fat;
	}
	return deChain;
}

DWORD getWorkbookDEOffset(DWORD* deChain, DWORD countDESec, char* buf, DWORD secSize) {
	DWORD offset = 0;
	for (int i = 0; i < countDESec; i++) {
		char* deSec = new char[secSize];
		ReadSector(deSec, buf, deChain[i], secSize);
		for (int j = 0; j < 4; j++) {
			DirectoryEntry de;
			memcpy(&de, &deSec[j * 128], 128);
			if (u16string(de.dirName) == u"Workbook") {
				offset = (deChain[i] + 1) * secSize + 128 * j;
				wcout << L"������� ����� �������!" << endl;
				return offset;
			}
		}
		delete[] deSec;
	}
	if (offset == 0) {
		wcout << L"������� ����� �� �������! ����������� ������!" << endl;
		exit(1);
	}
	return offset;
}

char* unpackWBSC(DWORD* wkbkSC, DWORD wbSSize, DWORD secSize, char* buf) {
	char* unp = new char[wbSSize * secSize];
	DWORD offset = 0;
	for (DWORD i = 0; i < wbSSize; i++) {
		ReadSector(&unp[offset], buf, wkbkSC[i], secSize);
		offset += secSize;
	}
	return unp;
}

vector<BoundSheet8> getbs8(char* Workbook, DWORD wbSSize, DWORD secSize, DWORD* outOffset) {
	vector<BoundSheet8> bs8;
	DWORD offset = 0;
	bool f = false;
	while (offset < wbSSize * secSize) {
		RecordHead rh;
		BoundSheet8 bs;
		memcpy(&rh, &Workbook[offset], 4);
		if (rh.type == 133 and f) {
			offset += 4;
			memcpy(&bs, &Workbook[offset], 8);
			offset += 8;
			DWORD strEnd = offset + rh.size - 8;
#ifdef _WIN32
			if (bs.name.fhighByte == 0) {
				while (offset < strEnd) {
					bs.name.rgb.push_back(Workbook[offset]);
					offset += 1;
					bs.name.rgb.push_back(0);
				}
			}
			else {
				bs.name.rgb.insert(bs.name.rgb.end(), &Workbook[offset], &Workbook[offset + rh.size - 8]);
				offset += rh.size - 8;
			}
#endif // _WIN32
#ifdef __linux__
			if (bs.name.fhighByte == 0) {
				while (offset < strEnd) {
					bs.name.rgb.push_back(Workbook[offset]);
					offset += 1;
					bs.name.rgb.push_back(0);
					bs.name.rgb.push_back(0);
					bs.name.rgb.push_back(0);
				}
			}
			else {
				while (offset < strEnd) {
					bs.name.rgb.push_back(Workbook[offset]);
					offset += 1;
					bs.name.rgb.push_back(Workbook[offset]);
					offset += 1;
					bs.name.rgb.push_back(0);
					bs.name.rgb.push_back(0);
				}
			}
#endif // __linux__
			bs8.push_back(bs);
		}
		else if (rh.type == 133 and !f) {
			f = true;
			offset += 4;
			memcpy(&bs, &Workbook[offset], 8);
			offset += 8;
			DWORD strEnd = offset + rh.size - 8;
#ifdef _WIN32
			if (bs.name.fhighByte == 0) {
				while (offset < strEnd) {
					bs.name.rgb.push_back(Workbook[offset]);
					offset += 1;
					bs.name.rgb.push_back(0);
				}
			}
			else {
				bs.name.rgb.insert(bs.name.rgb.end(), &Workbook[offset], &Workbook[offset + rh.size - 8]);
				offset += rh.size - 8;
			}
#endif // _WIN32
#ifdef __linux__
			if (bs.name.fhighByte == 0) {
				while (offset < strEnd) {
					bs.name.rgb.push_back(Workbook[offset]);
					offset += 1;
					bs.name.rgb.push_back(0);
					bs.name.rgb.push_back(0);
					bs.name.rgb.push_back(0);
				}
			}
			else {
				while (offset < strEnd) {
					bs.name.rgb.push_back(Workbook[offset]);
					offset += 1;
					bs.name.rgb.push_back(Workbook[offset]);
					offset += 1;
					bs.name.rgb.push_back(0);
					bs.name.rgb.push_back(0);
				}
			}
#endif // __linux__
			bs8.push_back(bs);
		}
		else if (rh.type != 133 and f) {
			break;
		}
		else {
			offset += 4 + rh.size;
		}
	}
	*outOffset = offset;
	return bs8;
}

inline void printSheetName(vector<BoundSheet8>* bsv, DWORD i) {
	wcout << wstring((const wchar_t*)(*bsv)[i].name.rgb.data(), (*bsv)[i].name.cch);
}

SST* getSST(char* Workbook, DWORD wbSSize, DWORD secSize, DWORD offset) {
	SST* sst = new SST;
	RecordHead rh;
	while (offset < wbSSize * secSize) {
		memcpy(&rh, &Workbook[offset], 4);
		if (rh.type == 252) {
			offset += 4;
			memcpy(sst, &Workbook[offset], 8);
			offset += 8;
			break;
		}
		else {
			offset += 4 + rh.size;
		}
	}
	DWORD sstEnd = offset + rh.size - 8;
	while (offset < sstEnd) {
		XLURES str;
		memcpy(&str, &Workbook[offset], 3);
		offset += 3;
		if (str.fRichSt == 1) {
			memcpy(&(str.cRun), &Workbook[offset], 2);
			offset += 2;
		}
		if (str.fExtSt == 1) {
			memcpy(&(str.cbExtRst), &Workbook[offset], 4);
			offset += 4;
		}
		DWORD strEnd = offset + str.cch * (str.fhighByte + 1);
		if (strEnd > sstEnd) {
			strEnd = sstEnd;
		}
#ifdef __linux__
		if (str.fhighByte == 1) {
			while (offset < strEnd) {
				str.rgb.push_back(Workbook[offset]);
				offset += 1;
				str.rgb.push_back(Workbook[offset]);
				offset += 1;
				str.rgb.push_back(0);
				str.rgb.push_back(0);
			}
		}
		else {
			while (offset < strEnd) {
				str.rgb.push_back(Workbook[offset]);
				offset += 1;
				str.rgb.push_back(0);
				str.rgb.push_back(0);
				str.rgb.push_back(0);
			}
		}
#else
		if (str.fhighByte == 0) {
			while (offset < strEnd) {
				str.rgb.push_back(Workbook[offset]);
				offset += 1;
				str.rgb.push_back(0);
			}
		}
		else {
			str.rgb.insert(str.rgb.end(), &Workbook[offset], &Workbook[strEnd]);
		}
#endif
		offset = strEnd + str.cRun * 4 + str.cbExtRst;
		sst->strings.push_back(str);
	}
	if (rh.size == 8224) {
		while (offset < wbSSize * secSize) {
			memcpy(&rh, &Workbook[offset], 4);
			if (rh.type == 60) {
				offset += 4;
				DWORD conOffset = offset;
				if (sst->strings.back().rgb.size() < (sst->strings.back().cch * sizeof(wchar_t))) {
					DWORD strEnd = offset + 1 + (sst->strings.back().cch - (sst->strings.back().rgb.size()) / sizeof(wchar_t)) * (Workbook[offset] + 1);
#ifdef __linux__
					if (Workbook[offset] == 1) {
						offset += 1;
						while (offset < strEnd) {
							sst->strings.back().rgb.push_back(Workbook[offset]);
							offset += 1;
							sst->strings.back().rgb.push_back(Workbook[offset]);
							offset += 1;
							sst->strings.back().rgb.push_back(0);
							sst->strings.back().rgb.push_back(0);
						}
					}
					else {
						offset += 1;
						while (offset < strEnd) {
							sst->strings.back().rgb.push_back(Workbook[offset]);
							offset += 1;
							sst->strings.back().rgb.push_back(0);
							sst->strings.back().rgb.push_back(0);
							sst->strings.back().rgb.push_back(0);
						}
					}
#else
					if (Workbook[offset] == 0) {
						offset += 1;
						while (offset < strEnd) {
							sst->strings.back().rgb.push_back(Workbook[offset]);
							offset += 1;
							sst->strings.back().rgb.push_back(0);
						}
					}
					else {
						sst->strings.back().rgb.insert(sst->strings.back().rgb.end(), &Workbook[offset + 1], &Workbook[strEnd]);
					}
#endif
					offset = strEnd + sst->strings.back().cRun * 4 + sst->strings.back().cbExtRst;
				}
				DWORD conEnd = conOffset + rh.size;
				while (offset < conEnd) {
					XLURES str;
					memcpy(&str, &Workbook[offset], 3);
					offset += 3;
					if (str.fRichSt == 1) {
						memcpy(&(str.cRun), &Workbook[offset], 2);
						offset += 2;
					}
					if (str.fExtSt == 1) {
						memcpy(&(str.cbExtRst), &Workbook[offset], 4);
						offset += 4;
					}
					DWORD strEnd = offset + str.cch * (str.fhighByte + 1);
					if (strEnd > conEnd) {
						strEnd = conEnd;
					}
#ifdef __linux__
					if (str.fhighByte == 1) {
						while (offset < strEnd) {
							str.rgb.push_back(Workbook[offset]);
							offset += 1;
							str.rgb.push_back(Workbook[offset]);
							offset += 1;
							str.rgb.push_back(0);
							str.rgb.push_back(0);
						}
					}
					else {
						while (offset < strEnd) {
							str.rgb.push_back(Workbook[offset]);
							offset += 1;
							str.rgb.push_back(0);
							str.rgb.push_back(0);
							str.rgb.push_back(0);
						}
					}
#else
					if (str.fhighByte == 0) {
						while (offset < strEnd) {
							str.rgb.push_back(Workbook[offset]);
							offset += 1;
							str.rgb.push_back(0);
						}
					}
					else {
						str.rgb.insert(str.rgb.end(), &Workbook[offset], &Workbook[strEnd]);
					}
#endif
					offset = strEnd + str.cRun * 4 + str.cbExtRst;
					sst->strings.push_back(str);
				}
			}
			else {
				break;
			}
		}

	}
	return sst;
}

inline void printXLURES(SST* sst, DWORD i) {
	wcout << wstring((const wchar_t*)sst->strings[i].rgb.data(), sst->strings[i].cch);
}

void printCell(SST* sst, LabelSst* lsst) {
	wcout << L"������ " << getCellIndex(lsst->cell.col + 1, lsst->cell.rw + 1) << L": ";
	printXLURES(sst, lsst->isst);
	wcout << endl;
}


#endif