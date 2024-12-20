#include <iostream>
#include <Windows.h>
#include <bit>
#include <vector>


using namespace std;

#pragma pack(push, 1)

/**
 * @struct CFHeader
 * @brief Представляет заголовок Compound File (CF), используемый в формате Compound File Binary Format (CFBF).
 *
 * Структура CFHeader содержит метаданные Compound File, включая информацию об управлении секторами,
 * версионировании и расположении основных потоков данных.
 */

struct CFHeader{
	BYTE Siganture[8]; ///< Сигнатура файла, используемая для идентификации его формата.
	GUID CLSID; ///< Классный идентификатор (CLSID), связанный с объектом хранения.
	WORD minorVersion; ///< Минорная версия формата Compound File.
	WORD majorVersion; ///< Основная версия формата Compound File.
	WORD byteOrder; ///< Идентификатор порядка байтов
	WORD sectorShift; ///< Размер сектора.
	WORD miniSectorShift; ///< Размер мини-сектора.
	BYTE reserved[6]; ///< Зарезервированные байты.
	DWORD numOfDirSectors; ///< Количество директорных секторов.
	DWORD numOfFATSectors; ///< Количество FAT-секторов.
	DWORD firstDirSecLoc; ///< Начальный сектор директории.
	DWORD TSN; ///< Номер подписи транзакции.
	DWORD miniStreamCUSize; ///< Минимальный поток.
	DWORD firstMiniFATSecLoc; ///< Первый сектор MiniFAT.
	DWORD numOfMiniFATSectors; ///< Количество MiniFAT-секторов.
	DWORD firstDIFATSecLoc; ///< Первый сектор DIFAT.
	DWORD numOfDIFATSectors; ///< Количество DIFAT-секторов.
	DWORD DIFAT[109]; ///< DIFAT-таблица.
};

/**
 * @struct DirectoryEntry
 * @brief Представляет запись каталога в формате Compound File Binary Format (CFBF).
 *
 * Структура используется для описания объектов, содержащихся в документе Compound File, включая их имя,
 * тип, связи с другими объектами, а также метаданные, такие как время создания и изменения.
 */

struct DirectoryEntry {
	WCHAR dirName[32]; ///< Имя директории.
	WORD nameLength; ///< Длина имени.
	BYTE objType; ///< Тип объекта.
	BYTE colorFlag; ///< Флаг цвета.
	DWORD leftSibID; ///< Левый дочерний элемент.
	DWORD rightSibID; ///< Правый дочерний элемент.
	DWORD childID; ///< ID дочернего элемента.
	GUID CLSID; ///< CLSID объекта.
	DWORD stateBits; ///< Состояние объекта.
	ULONGLONG creationTime; ///< Время создания.
	ULONGLONG modifiedTime; ///< Время изменения.
	DWORD startingSecLoc; ///< Начальный сектор.
	ULONGLONG streamSize; ///< Размер потока.
};

/**
 * @struct RecordHead
 * @brief Заголовок записи в структуре данных документа Compound File Binary Format (CFBF).
 *
 * Используется для описания общей структуры записи, включая ее тип и размер.
 */

struct RecordHead {
	WORD type; ///< Тип записи.
	WORD size; ///< Размер записи
};

/**
 * @struct WsBool
 * @brief Флаги настроек рабочего листа в формате Excel.
 *
 * Структура содержит битовые поля, определяющие различные параметры и поведение рабочего листа.
 */

struct WsBool {
	BYTE fShowAutoBreaks : 1; ///< Флаг отображения автоматических разрывов страниц.
	BYTE reserved1 : 3; ///< Зарезервировано.
	BYTE fDialog : 1; ///< Флаг диалогового окна.
	BYTE fApplyStyles : 1; ///< Флаг применения стилей при печати.
	BYTE fRowSumsBelow : 1; ///< Флаг сумм по строкам снизу.
	BYTE fColSumsRight : 1; ///< Флаг сумм по столбцам справа.
	BYTE fFitToPage : 1; ///< Флаг подгонки содержимого страницы к размеру при печати.
	BYTE reserved2 : 1; ///< Зарезервировано.
	BYTE unused : 2; ///< Неиспользуемые биты.
	BYTE fSyncHoriz : 1; ///< Флаг синхронизации горизонтальной прокрутки.
	BYTE fSyncVert : 1; ///< Флаг синхронизации вертикальной прокрутки.
	BYTE fAltExprEval : 1; ///< Флаг альтернативной оценки выражений.
	BYTE fAltFormulaEntry : 1; ///< Флаг альтернативного ввода формул.
};

/**
 * @struct SXLUS
 * @brief Структура для представления строки в формате Excel.
 *
 * Данная структура используется для хранения строки с информацией о длине и параметрами,
 * такими как наличие символов верхнего байта (широкие символы).
 */

struct SXLUS {
	BYTE cch; ///< Длина строки.
	BYTE fHighByte : 1; ///< Флаг, указывающий, содержит ли строка символы верхнего байта.
	BYTE reserved : 7; ///< Зарезервировано.
	vector<BYTE> rgb; ///< Вектор байтов, представляющий строку в кодировке, зависящей от значения `fHighByte`.
};

/**
 * @struct BoundSheet8
 * @brief Структура для представления метаданных о листе в файле Excel.
 *
 * Эта структура описывает информацию о листе, такую как его позиция в потоке, тип и имя.
 */

struct BoundSheet8 {
	DWORD pos; ///< Позиция листа в потоке Workbook.
	BYTE trash; ///< Неиспользуемое или резервное поле.
	BYTE dt; ///< Тип листа.
	SXLUS name; ///< Имя листа.
};

/**
 * @struct XLURES
 * @brief Представляет строку в формате BIFF (Binary Interchange File Format) Excel.
 *
 * Структура используется для хранения данных строки, включая ее длину, атрибуты и содержимое.
 */

struct XLURES {
	WORD cch; ///< Количество символов в строке.
	BYTE fHighByte : 1; ///< Флаг, указывающий, содержит ли строка символы Unicode.
	BYTE reserved1 : 1; ///< Зарезервированное поле.
	BYTE fExtSt : 1; ///< Флаг наличия расширенных данных строки.
	BYTE fRichSt : 1; ///< Флаг наличия данных форматирования (rich text).
	BYTE reserved2 : 4; ///< Зарезервированное поле.
	WORD cRun = 0; ///< Количество пар форматирования rich text.
	DWORD cbExtRst = 0; ///< Размер расширенных данных строки (в байтах).
	vector<BYTE> rgb; ///< Содержимое строки.
};

/**
 * @struct SST
 * @brief Структура, представляющая таблицу строк (String Table) в формате BIFF (Binary Interchange File Format) для Excel.
 *
 * Таблица строк используется в Excel для хранения всех строковых данных (например, текстовые метки, надписи в ячейках, и т.д.),
 * которые могут быть повторно использованы в различных частях документа.
 */

struct SST {
	DWORD cstTotal; ///< Общее количество строк в таблице.
	DWORD cstUnique; ///< Количество уникальных строк в таблице.
	vector<XLURES> strings; ///< Вектор, содержащий уникальные строки таблицы.
};

#pragma pack(pop)

/**
 * @brief Сравнивает два массива байтов на равенство.
 *
 * Функция проверяет, идентичны ли два массива байтов указанного размера. Массивы считаются одинаковыми, если все элементы в них совпадают по порядку.
 *
 * @param arr1 Указатель на первый массив байтов.
 * @param arr2 Указатель на второй массив байтов.
 * @param size Размер массивов, который должен быть одинаковым для обоих массивов.
 *
 * @return `true`, если массивы идентичны, `false`, если хотя бы один элемент в массивах не совпадает.
 * @code
 bool isEqualArr(BYTE* arr1, BYTE* arr2, DWORD size) {
	bool f = true;
	for (int i = 0; i < size; i++) {
		if (arr1[i] != arr2[i]) {
			f = false;
			break;
		}
	}
	return f;
 }
 * @endcode
 */

bool isEqualArr(BYTE* arr1, BYTE* arr2, DWORD size) {
	bool f = true;
	for (int i = 0; i < size; i++) {
		if (arr1[i] != arr2[i]) {
			f = false;
			break;
		}
	}
	return f;
}

/**
 * @brief Читает данные сектора из буфера в указанную область памяти.
 *
 * Функция копирует данные сектора из буфера в указанную область памяти.
 * Сектор определяется его номером и размером, который задается в параметре.
 *
 * @param dst Указатель на область памяти, в которую будут записаны данные сектора.
 * @param buf Указатель на буфер, содержащий данные всех секторов.
 * @param secNum Номер сектора, данные которого нужно извлечь.
 * @param size Размер сектора в байтах.
 * @code
 inline void ReadSector(void* dst, BYTE* buf, DWORD secNum, DWORD size) {
	memcpy(dst, &buf[(secNum + 1) * size], size);
 }
 * @endcode
 */

inline void ReadSector(void* dst, BYTE* buf, DWORD secNum, DWORD size) {
	memcpy(dst, &buf[(secNum + 1) * size], size);
}

/**
 * @brief Извлекает цепочку DIFAT секторов и сохраняет их в массив.
 *
 * Функция извлекает цепочку секторов DIFAT (FAT для больших файлов) и сохраняет их в указанный массив.
 *
 * @param dst Указатель на массив, в который будут записаны номера секторов DIFAT.
 * @param cfh Указатель на структуру заголовка CF (Compound File), которая содержит информацию о секторах DIFAT.
 * @param buf Указатель на буфер, содержащий все данные файловой системы.
 * @param secSize Размер сектора в байтах.
 * @code
 void getDIFATChain(DWORD* dst, CFHeader* cfh, BYTE* buf, DWORD secSize) {
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
 * @endcode
 */

void getDIFATChain(DWORD* dst, CFHeader* cfh, BYTE* buf, DWORD secSize) {
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

/**
 * @brief Извлекает цепочку FAT секторов и сохраняет их в массив.
 *
 * Функция извлекает цепочку секторов FAT, начиная с данных, которые содержатся в DIFAT (сектора для FAT),
 * и сохраняет их в указанный массив.
 *
 * Функция сначала копирует секторы FAT, указанные в структуре `cfh->DIFAT`, в массив `dst`. Если количество
 * секторов FAT больше 109, то дополнительная информация о FAT считывается из секторов DIFAT.
 *
 * @param dst Указатель на массив, в который будут записаны номера секторов FAT.
 * @param difCh Массив, содержащий номера секторов DIFAT.
 * @param cfh Указатель на структуру заголовка CF (Compound File), содержащую информацию о секторах DIFAT и FAT.
 * @param buf Указатель на буфер, содержащий все данные файловой системы.
 * @param secSize Размер сектора в байтах.
 * @code
 void getFATChain(DWORD* dst, DWORD* difCh, CFHeader* cfh, BYTE* buf, DWORD secSize) {
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
 * @endcode
 */

void getFATChain(DWORD* dst, DWORD* difCh, CFHeader* cfh, BYTE* buf, DWORD secSize) {
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

/**
 * @brief Извлекает цепочку секторов книги из FAT и сохраняет их в массив.
 *
 * Функция извлекает цепочку секторов рабочей книги, используя информацию из FAT,
 * и сохраняет её в массив. Каждый сектор FAT указывает на следующий сектор в цепочке,
 * и эта информация используется для восстановления последовательности секторов,
 * содержащих данные книги.
 *
 * Для каждой позиции в цепочке, функция рассчитывает, какой сектор FAT нужно использовать,
 * и читает его, чтобы найти адрес следующего сектора. Этот процесс повторяется до тех пор,
 * пока не будет извлечено нужное количество секторов.
 *
 * @param dst Массив, в который будут записаны номера секторов рабочей книги.
 * @param fatCh Массив с номерами секторов FAT.
 * @param de Указатель на структуру DirectoryEntry, содержащую информацию о начале цепочки секторов.
 * @param chainSize Размер цепочки секторов, которую нужно извлечь.
 * @param secSize Размер сектора в байтах.
 * @param buf Указатель на буфер, содержащий все данные файловой системы.
 * @code
 void getWorkbookChain(DWORD* dst, DWORD* fatCh, DirectoryEntry* de, DWORD chainSize, DWORD secSize, BYTE* buf) {
	dst[0] = de->startingSecLoc;
	for (int i = 1; i < chainSize; i++) {
        DWORD curFATSecIndex = dst[i - 1] / (secSize / 4);
		DWORD* fat = new DWORD[secSize / 4];
		ReadSector(fat, buf, fatCh[curFATSecIndex], secSize);
		dst[i] = fat[dst[i - 1] % (secSize / 4)];
		delete[] fat;
	}
 }
 * @endcode
 */

void getWorkbookChain(DWORD* dst, DWORD* fatCh, DirectoryEntry* de, DWORD chainSize, DWORD secSize, BYTE* buf) {
	dst[0] = de->startingSecLoc;
	for (int i = 1; i < chainSize; i++) {
		DWORD curFATSecIndex = dst[i - 1] / (secSize / 4);
		DWORD* fat = new DWORD[secSize / 4];
		ReadSector(fat, buf, fatCh[curFATSecIndex], secSize);
		dst[i] = fat[dst[i - 1] % (secSize / 4)];
		delete[] fat;
	}
}

/**
 * @brief Открывает файл для чтения.
 *
 * Функция открывает файл с указанным именем для чтения с использованием
 * Windows API. Если файл успешно открыт, возвращается дескриптор файла.
 * В случае ошибки при открытии файла выводится сообщение об ошибке,
 * и программа завершает выполнение.
 *
 * @param filename Указатель на строку, содержащую имя файла, который необходимо открыть.
 *
 * @return Возвращает дескриптор открытого файла в случае успеха.
 * @code
 HANDLE openFile(const wchar_t* filename) {
	HANDLE fileHandle = CreateFile(filename, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (fileHandle == INVALID_HANDLE_VALUE) {
		wcout << L"Ошибка открытия файла!" << endl;
		CloseHandle(fileHandle);
		exit(1);
	}
	return fileHandle;
 }
 * @endcode
 */

HANDLE openFile(const wchar_t* filename) {
	HANDLE fileHandle = CreateFile(filename, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (fileHandle == INVALID_HANDLE_VALUE) {
		wcout << L"Ошибка открытия файла!" << endl;
		CloseHandle(fileHandle);
		exit(1);
	}
	return fileHandle;
}

/**
 * @brief Читает данные из файла в буфер.
 *
 * Функция читает всё содержимое указанного файла в динамически выделенный буфер
 * и возвращает указатель на этот буфер. Если чтение файла завершается с ошибкой,
 * выводится сообщение об ошибке, и программа завершает выполнение.
 *
 * @param fileHandle Дескриптор открытого файла, из которого нужно считать данные.
 *
 * @return Возвращает указатель на буфер с прочитанными данными из файла.
 * @code
 BYTE* getData(HANDLE fileHandle) {
	DWORD fileSize = GetFileSize(fileHandle, NULL);
	DWORD bytesRead;
	BYTE* buf = new BYTE[fileSize];
	BOOL readOK = ReadFile(fileHandle, buf, fileSize, &bytesRead, NULL);
	if (!readOK) {
		wcout << L"Ошибка чтения файла!" << endl;
		CloseHandle(fileHandle);
		exit(1);
	}
	CloseHandle(fileHandle);
	return buf;
 }
 * @endcode
 */

BYTE* getData(HANDLE fileHandle) {
	DWORD fileSize = GetFileSize(fileHandle, NULL);
	DWORD bytesRead;
	BYTE* buf = new BYTE[fileSize];
	BOOL readOK = ReadFile(fileHandle, buf, fileSize, &bytesRead, NULL);
	if (!readOK) {
		wcout << L"Ошибка чтения файла!" << endl;
		CloseHandle(fileHandle);
		exit(1);
	}
	CloseHandle(fileHandle);
	return buf;
}

/**
 * @brief Читает заголовок из буфера и возвращает указатель на структуру CFHeader.
 *
 * Функция выделяет память для структуры `CFHeader`, копирует данные из буфера
 * в структуру и возвращает указатель на эту структуру.
 *
 * @param buf Указатель на буфер, содержащий данные заголовка.
 *
 * @return Указатель на структуру CFHeader, заполненную данными из буфера.
 * @code
 CFHeader* readHeader(BYTE* buf) {
	CFHeader* cfh = new CFHeader;
	memcpy(cfh, &buf[0], sizeof(CFHeader));
	return cfh;
 }
 * @endcode
 */

CFHeader* readHeader(BYTE* buf) {
	CFHeader* cfh = new CFHeader;
	memcpy(cfh, &buf[0], sizeof(CFHeader));
	return cfh;
}

/**
 * @brief Проверяет подпись заголовка файла для подтверждения формата Compound File.
 *
 * Функция проверяет, соответствует ли подпись заголовка переданного объекта `CFHeader`
 * заранее известной подписи для формата Compound File. Если подпись неверная, программа завершится с ошибкой.
 * Если подпись верная, выводится сообщение о подтверждении формата.
 *
 * @param cfh Указатель на структуру CFHeader, содержащую подпись заголовка файла.
 *
 * @return Нет возвращаемого значения. В случае ошибки программа завершится.
 * @code
 void checkSig(CFHeader* cfh) {
	BYTE trueHeaderSig[] = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
	if (!isEqualArr(trueHeaderSig, cfh->Siganture, 8)) {
		cout << "File isn't a Compound File!" << endl;;
		exit(1);
	}
	else {
		cout << "Compound file confirmed!" << endl;
	}
 }
 * @endcode
 */

void checkSig(CFHeader* cfh) {
	BYTE trueHeaderSig[] = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
	if (!isEqualArr(trueHeaderSig, cfh->Siganture, 8)) {
		cout << "File isn't a Compound File!" << endl;;
		exit(1);
	}
	else {
		cout << "Compound file confirmed!" << endl;
	}
}

/**
 * @brief Подсчитывает количество секторов, связанных с цепочкой Directory Entry (DE).
 *
 * Функция вычисляет количество секторов в цепочке DE, начиная с указанного первого сектора.
 * Она использует FAT для поиска следующего сектора в цепочке, пока не встретит специальный маркер конца цепочки (0xFFFFFFFE).
 *
 * @param fstDELoc Номер первого сектора в цепочке Directory Entry.
 * @param buf Указатель на буфер с данными файла.
 * @param fatCh Указатель на массив с номерами секторов FAT.
 * @param secSize Размер сектора в байтах.
 *
 * @return Количество секторов в цепочке Directory Entry.
 * @code
 DWORD countDESectors(DWORD fstDELoc, BYTE* buf, DWORD* fatCh, DWORD secSize) {
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
 * @endcode
 */

DWORD countDESectors(DWORD fstDELoc, BYTE* buf, DWORD* fatCh, DWORD secSize) {
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

 /**
 * @brief Создает цепочку секторов Directory Entry (DE) из FAT.
 *
 * Эта функция формирует массив секторов, представляющий цепочку Directory Entry (DE), начиная с указанного начального сектора.
 * Память под массив выделяется динамически, и её необходимо освободить с помощью `delete[]` после использования.
 *
 * @param fstDELoc Номер начального сектора в цепочке Directory Entry.
 * @param buf Указатель на буфер с данными Compound File.
 * @param fatCh Указатель на массив секторов FAT, определяющий структуру цепочек.
 * @param secSize Размер сектора в байтах.
 * @param k Количество секторов в цепочке Directory Entry.
 *
 * @return Указатель на массив номеров секторов цепочки Directory Entry.
 * @code
 DWORD* getDEChain(DWORD fstDELoc, BYTE* buf, DWORD* fatCh, DWORD secSize, DWORD k) {
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
 * @endcode
 */

DWORD* getDEChain(DWORD fstDELoc, BYTE* buf, DWORD* fatCh, DWORD secSize, DWORD k) {
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

/**
 * @brief Определяет смещение Directory Entry (DE) объекта Workbook в файле.
 *
 * Функция сканирует цепочку секторов Directory Entry (DE), чтобы найти объект с именем "Workbook"
 * и возвращает его смещение относительно начала файла.
 *
 * @param deChain Указатель на массив номеров секторов, формирующих цепочку Directory Entry.
 * @param countDESec Количество секторов в цепочке Directory Entry.
 * @param buf Указатель на буфер данных Compound File.
 * @param secSize Размер сектора в байтах.
 *
 * @return Смещение Directory Entry объекта Workbook. Если объект не найден, программа завершает выполнение с ошибкой.
 * @code
 DWORD getWorkbookDEOffset(DWORD* deChain, DWORD countDESec, BYTE* buf, DWORD secSize) {
	DWORD offset = 0;
	for (int i = 0; i < countDESec; i++) {
		BYTE* deSec = new BYTE[secSize];
		ReadSector(deSec, buf, deChain[i], secSize);
		for (int j = 0; j < 4; j++) {
			DirectoryEntry de;
			memcpy(&de, &deSec[j * 128], 128);
			if (wstring(de.dirName) == L"Workbook") {
				offset = (deChain[i] + 1) * secSize + 128 * j;
				cout << "Workbook found!" << endl;
				return offset;
			}
		}
		delete[] deSec;
	}
	if (offset == 0) {
		cout << "Workbook wasn't found! Closing..." << endl;
		exit(1);
	}
	return offset;
 }
 * @endcode
 */

DWORD getWorkbookDEOffset(DWORD* deChain, DWORD countDESec, BYTE* buf, DWORD secSize) {
	DWORD offset = 0;
	for (int i = 0; i < countDESec; i++) {
		BYTE* deSec = new BYTE[secSize];
		ReadSector(deSec, buf, deChain[i], secSize);
		for (int j = 0; j < 4; j++) {
			DirectoryEntry de;
			memcpy(&de, &deSec[j * 128], 128);
			if (wstring(de.dirName) == L"Workbook") {
				offset = (deChain[i] + 1) * secSize + 128 * j;
				cout << "Workbook found!" << endl;
				return offset;
			}
		}
		delete[] deSec;
	}
	if (offset == 0) {
		cout << "Workbook wasn't found! Closing..." << endl;
		exit(1);
	}
	return offset;
}

/**
 * @brief Распаковывает цепочку секторов Workbook в единый непрерывный буфер.
 *
 * Функция читает данные из цепочки секторов, принадлежащих Workbook,
 * и объединяет их в непрерывный массив байтов.
 *
 * @param wkbkSC Указатель на массив номеров секторов, составляющих цепочку Workbook.
 * @param wbSSize Количество секторов в цепочке Workbook.
 * @param secSize Размер сектора в байтах.
 * @param buf Указатель на буфер данных Compound File.
 *
 * @return Указатель на новый буфер, содержащий данные из всех секторов Workbook.
 * @code
 BYTE* unpackWBSC(DWORD* wkbkSC, DWORD wbSSize, DWORD secSize, BYTE* buf) {
	BYTE* unp = new BYTE[wbSSize * secSize];
	DWORD offset = 0;
	for (DWORD i = 0; i < wbSSize; i++) {
		ReadSector(&unp[offset], buf, wkbkSC[i], secSize);
		offset += secSize;
	}
	return unp;
 }
 * @endcode
 */

BYTE* unpackWBSC(DWORD* wkbkSC, DWORD wbSSize, DWORD secSize, BYTE* buf) {
	BYTE* unp = new BYTE[wbSSize * secSize];
	DWORD offset = 0;
	for (DWORD i = 0; i < wbSSize; i++) {
		ReadSector(&unp[offset], buf, wkbkSC[i], secSize);
		offset += secSize;
	}
	return unp;
}

/**
 * @brief Извлекает массив объектов BoundSheet8 из данных Workbook.
 *
 * Функция парсит содержимое Workbook и извлекает записи типа BoundSheet8,
 * которые описывают листы электронной таблицы. Она возвращает вектор с объектами BoundSheet8
 * и смещение, на котором остановился парсер.
 *
 * @param Workbook Указатель на буфер данных Workbook.
 * @param wbSSize Размер Workbook в секторах.
 * @param secSize Размер сектора в байтах.
 * @param outOffset Указатель на переменную, в которую будет записано итоговое смещение парсинга.
 *
 * @return Вектор объектов BoundSheet8, содержащих информацию о листах.
 * @code
 vector<BoundSheet8> getbs8(BYTE* Workbook, DWORD wbSSize, DWORD secSize, DWORD* outOffset) {
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
			bs.name.rgb.insert(bs.name.rgb.end(), &Workbook[offset], &Workbook[offset + rh.size - 8]);
			bs8.push_back(bs);
			offset += rh.size - 8;
		}
		else if (rh.type == 133 and !f) {
			f = true;
			offset += 4;
			memcpy(&bs, &Workbook[offset], 8);
			offset += 8;
			bs.name.rgb.insert(bs.name.rgb.end(), &Workbook[offset], &Workbook[offset + rh.size - 8]);
			bs8.push_back(bs);
			offset += rh.size - 8;
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
 * @endcode
 */

vector<BoundSheet8> getbs8(BYTE* Workbook, DWORD wbSSize, DWORD secSize, DWORD* outOffset) {
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
			bs.name.rgb.insert(bs.name.rgb.end(), &Workbook[offset], &Workbook[offset + rh.size - 8]);
			bs8.push_back(bs);
			offset += rh.size - 8;
		}
		else if (rh.type == 133 and !f) {
			f = true;
			offset += 4;
			memcpy(&bs, &Workbook[offset], 8);
			offset += 8;
			bs.name.rgb.insert(bs.name.rgb.end(), &Workbook[offset], &Workbook[offset + rh.size - 8]);
			bs8.push_back(bs);
			offset += rh.size - 8;
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

/**
 * @brief Выводит имя листа из вектора BoundSheet8.
 *
 * Функция принимает вектор объектов BoundSheet8, индекс листа и выводит его имя
 * на консоль. Поддерживает как ANSI, так и Unicode кодировки.
 *
 * @param bsv Вектор объектов BoundSheet8, содержащих метаданные листов.
 * @param i Индекс листа, имя которого требуется вывести.
 * @code
 void printSheetName(vector<BoundSheet8> bsv, DWORD i) {
	if (bsv[i].name.fHighByte == 0) {
		cout << string((const char*)bsv[i].name.rgb.data()) << endl;
	}
	else {
		wcout << wstring((const wchar_t*)bsv[i].name.rgb.data()) << endl;
	}
 }
 * @endcode
 */

void printSheetName(vector<BoundSheet8> bsv, DWORD i) {
	if (bsv[i].name.fHighByte == 0) {
		cout << string((const char*)bsv[i].name.rgb.data()) << endl;
	}
	else {
		wcout << wstring((const wchar_t*)bsv[i].name.rgb.data()) << endl;
	}
}

/**
 * @brief Извлекает таблицу общих строк (Shared String Table, SST) из Workbook.
 *
 * Функция считывает структуру SST из потока Workbook, включая строки с расширенными свойствами.
 * SST используется для хранения строковых данных, которые используются в ячейках листов Excel.
 *
 * @param Workbook Указатель на массив байтов, представляющий содержимое Workbook.
 * @param wbSSize Количество секторов в Workbook.
 * @param secSize Размер одного сектора в байтах.
 * @param offset Смещение, с которого начинается поиск SST в Workbook.
 * @return Указатель на структуру SST, содержащую строки из таблицы общих строк.
 * @code
 SST* getSST(BYTE* Workbook, DWORD wbSSize, DWORD secSize, DWORD offset) {
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
		DWORD strEnd = offset + str.cch * (str.fHighByte + 1);
		if (strEnd > sstEnd) {
			strEnd = sstEnd;
		}
		str.rgb.insert(str.rgb.end(), &Workbook[offset], &Workbook[strEnd]);
		offset = strEnd + str.cRun * 4 + str.cbExtRst;
		sst->strings.push_back(str);
	}
	if (rh.size == 8224) {
		while (offset < wbSSize * secSize) {
			memcpy(&rh, &Workbook[offset], 4);
			if (rh.type == 60) {
				offset += 4;
				DWORD conOffset = offset;
				if (sst->strings.back().rgb.size() < (sst->strings.back().cch * (sst->strings.back().fHighByte + 1))) {
					DWORD strEnd = offset + 1 + sst->strings.back().cch * (sst->strings.back().fHighByte + 1) - sst->strings.back().rgb.size();
					sst->strings.back().rgb.insert(sst->strings.back().rgb.end(), &Workbook[offset + 1], &Workbook[strEnd]);
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
					DWORD strEnd = offset + str.cch * (str.fHighByte + 1);
					if (strEnd > conEnd) {
						strEnd = conEnd;
					}
					str.rgb.insert(str.rgb.end(), &Workbook[offset], &Workbook[strEnd]);
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
 * @endcode
 */

SST* getSST(BYTE* Workbook, DWORD wbSSize, DWORD secSize, DWORD offset) {
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
		DWORD strEnd = offset + str.cch * (str.fHighByte + 1);
		if (strEnd > sstEnd) {
			strEnd = sstEnd;
		}
		str.rgb.insert(str.rgb.end(), &Workbook[offset], &Workbook[strEnd]);
		offset = strEnd + str.cRun * 4 + str.cbExtRst;
		sst->strings.push_back(str);
	}
	if (rh.size == 8224) {
		while (offset < wbSSize * secSize) {
			memcpy(&rh, &Workbook[offset], 4);
			if (rh.type == 60) {
				offset += 4;
				DWORD conOffset = offset;
				if (sst->strings.back().rgb.size() < (sst->strings.back().cch * (sst->strings.back().fHighByte + 1))) {
					DWORD strEnd = offset + 1 + sst->strings.back().cch * (sst->strings.back().fHighByte + 1) - sst->strings.back().rgb.size();
					sst->strings.back().rgb.insert(sst->strings.back().rgb.end(), &Workbook[offset + 1], &Workbook[strEnd]);
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
					DWORD strEnd = offset + str.cch * (str.fHighByte + 1);
					if (strEnd > conEnd) {
						strEnd = conEnd;
					}
					str.rgb.insert(str.rgb.end(), &Workbook[offset], &Workbook[strEnd]);
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

/**
 * @brief Печатает строку XLURES из таблицы SST.
 *
 * Функция принимает объект SST и индекс строки, после чего выводит строку на консоль.
 * Строка может быть представлена либо в однобайтовой (ASCII), либо в двухбайтовой (Unicode) кодировке.
 *
 * @param sst Указатель на объект SST, содержащий строки.
 * @param i Индекс строки в таблице SST.
 * @code
 void printXLURES(SST* sst, DWORD i) {
	XLURES str = sst->strings[i];
	if (str.fHighByte == 0) {
		cout << string((const char*)str.rgb.data(), str.cch) << endl;
	}
	else {
		wcout << wstring((const wchar_t*)str.rgb.data(), str.cch) << endl;
	}
 }
 * @endcode
 */

void printXLURES(SST* sst, DWORD i) {
	XLURES str = sst->strings[i];
	if (str.fHighByte == 0) {
		cout << string((const char*)str.rgb.data(), str.cch) << endl;
	}
	else {
		wcout << wstring((const wchar_t*)str.rgb.data(), str.cch) << endl;
	}
}

int main() {

	setlocale(LC_ALL, "Russian");

	HANDLE fileHandle = openFile(L"6800.xls");
	BYTE* buf = getData(fileHandle);
	CFHeader* cfh = readHeader(buf);
	checkSig(cfh);
	DWORD sectorSize = pow(2, cfh->sectorShift);
	DWORD nds = cfh->numOfDIFATSectors ? cfh->numOfDIFATSectors : 1;
	DWORD* DIFATChain = new DWORD[nds];
	getDIFATChain(DIFATChain, cfh, buf, sectorSize);
	DWORD* FATChain = new DWORD[cfh->numOfFATSectors];
	getFATChain(FATChain, DIFATChain, cfh, buf, sectorSize);
	DWORD k = countDESectors(cfh->firstDirSecLoc, buf, FATChain, sectorSize);
	DWORD* deChain = getDEChain(cfh->firstDirSecLoc, buf, FATChain, sectorSize, k);
	DirectoryEntry de;
	DWORD deOffset = getWorkbookDEOffset(deChain, k, buf, sectorSize);
	memcpy(&de, &buf[deOffset], 128);
	DWORD wbSSize = ceil(double(de.streamSize) / sectorSize);
	DWORD* WorkbookSC = new DWORD[wbSSize];
	getWorkbookChain(WorkbookSC, FATChain, &de, wbSSize, sectorSize, buf);
	BYTE* Workbook = unpackWBSC(WorkbookSC, wbSSize, sectorSize, buf);
	delete[] buf;
	DWORD offset = 0;
	vector<BoundSheet8> bsv = getbs8(Workbook, wbSSize, sectorSize, &offset);
	SST* sst = getSST(Workbook, wbSSize, sectorSize, offset);
	printXLURES(sst, 385);







	delete[] Workbook;
	delete cfh;
	delete[] DIFATChain;
	delete[] FATChain;
	delete[] deChain;
	delete[] WorkbookSC;
	delete sst;

	return 0;
}
