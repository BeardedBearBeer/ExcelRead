#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stdarg.h>
#include "excel_read.h"

#define PATH_LENGTH 256
#define COMMAND_LENGTH 256
#define COPY_FILENAME_INS_STR "_copy"
#define START_TAG "<v>" /* XMLファイルの数字の開始タグ */
#define END_TAG "</v>" /* XMLファイルの数字の終了タグ */
#define STRING_EXPRESSION "t=\"s\""
#define MOVE_TO_ID 11
#define MOVE_TO_STRING_EXPRESSION 6

/*型判別用の構造体
enum type_of_number {
	TYPE_INT = 0,
	TYPE_DOUBLE,
	TYPE_FLOAT,
};

typedef struct intData {
	char id;
	int data;
};

typedef struct doubleData {
	char id;
	double data;
};

typedef struct floatData {
	char id;
	float data;
};
*/

/* プロトタイプ宣言--------------------------------------------------------------------------------------------------------------------------------------------------- */
static char *xlsx2zip(const char *path);
static void unzip(const char *zip_path);
static char *sheet_file_path_get(char *sheet, char *zip_path);
static void remove_expanded_data(char *zip_path);


/************************************************************
* 関数名：xlsx_nums_read
* 機能：指定したXLSXファイルのシートから指定された変数に値を格納する関数
* 引数： buf 値を格納したい変数のアドレスを指定
*		path XLSXファイルのパスを文字列で指定
*		sheet XLSXファイルのシート名を文字列で指定
 ************************************************************/
void xlsx_nums_read(void *buf, const char *path, char *sheet)
{
	FILE *fp;
	char *zip_path;
	int *intp = (int *)buf;
	/* double *doublep = (double *)buf; */ /* double型未実装 */
	char tmp[5000]; /* XMLファイルのデータ格納用 */
	char *scan; /* データ走査用ポインタ */
	int i = 0;

	if(!(strstr(path, ".xlsx")))
	{
		printf("Please specify XLSX format file as path.\n");
		exit(0);
	}

	zip_path = xlsx2zip(path);
	unzip(zip_path);

	/* 指定したシート名に対応するファイルのパスを取得 */
	sheet = sheet_file_path_get(sheet, zip_path);

	/* 取得したシートファイルを開いて全てのデータを tmp に格納 */
	if((fp = fopen(sheet, "r")) == NULL)
	{
		printf("File Not found.\n");
		exit(1);
	}
	else
	{
		while((fgets(tmp, 5000, fp)) != NULL);

		/* データ走査用ポインタに 開始タグ<v> からのアドレスを渡す */
		scan = strstr(tmp, START_TAG);
	}
	fclose(fp);

	do {
		/* 開始タグ<v> と 終了タグ</v> で囲まれている部分（数字）があれば */
		if((strstr(scan, START_TAG) && strstr(scan, END_TAG)))
		{
			/* セルに格納されているデータが文字列か数字かを判別 */
			if((strncmp((scan - MOVE_TO_STRING_EXPRESSION), STRING_EXPRESSION, strlen(STRING_EXPRESSION))))
			{
				/* 数字を数値に変換(double型)して引数で渡された配列に格納 */
				intp[i++] = atoi(scan + strlen(START_TAG));
				/* printf("number[%d] = %lf\n", i, doublep[i]); */
			}

			/* データ走査用ポインタを次の開始タグへ進める */
			scan = strstr(scan, END_TAG) + strlen(END_TAG);
			scan = strstr(scan, START_TAG);
		}
		else
		{
			remove_expanded_data(zip_path);
			printf("Numbers could not be found.\n");
			exit(1);
		}
	} while(scan);

	/* 作成されたZIPファイル、展開されたファイル・フォルダの削除 */
	remove_expanded_data(zip_path);
}


/************************************************************
* 関数名：xlsx_nums_read_from_cells
* 機能：指定したXLSXファイルのシート及びセルから指定された変数に値を格納する関数
* 引数： buf 値を格納したい変数のアドレスを指定
*		path XLSXファイルのパスを文字列で指定
*		sheet XLSXファイルのシート名を文字列で指定
*		cell_num 対象とするセルの個数を指定
*		... 対象とするセル名をカンマ区切りで指定
 ************************************************************/
void xlsx_nums_read_from_cells(void *buf, const char *path, char *sheet, int cell_num, ...)
{
	FILE *fp;
	char *zip_path;
	int *intp = (int *)buf;
	/* double *doublep = (double *)buf; */ /* double型未実装 */
	char tmp[5000]; /* XMLファイルのデータ格納用 */
	char *scan; /* データ走査用ポインタ */
	int i;
	va_list args;
	char **cells_name;

	if(!(strstr(path, ".xlsx")))
	{
		printf("Please specify XLSX format file as path.\n");
		exit(0);
	}

	/* 可変長引数に渡されたセル名の個数分メモリを確保 */
	cells_name = (char **)malloc(cell_num * sizeof(char *));
	memset(cells_name, NULL, cell_num * sizeof(cells_name));

	/* 可変長引数に渡されたセル名を取得 */
	va_start(args, cell_num);
	for(i = 0; i < cell_num; i++)
	{
		cells_name[i] = va_arg(args, char *);
		/* printf("%s\n", cells_name[i]); */
	}

	/* path に指定されたXLSXファイルのコピーファイルをZIP化 */
	zip_path = xlsx2zip(path);
	/* ZIP化したファイルを展開 */
	unzip(zip_path);

	/* 指定したシート名に対応するデータが記述されているファイルのパスを取得 */
	sheet = sheet_file_path_get(sheet, zip_path);

	/* 取得したシートファイルを開いて全てのデータを tmp に格納 */
	if((fp = fopen(sheet, "r")) == NULL)
	{
		printf("File Not found.\n");
		exit(1);
	}
	else
	{
		while((fgets(tmp, 5000, fp)) != NULL);
	}
	fclose(fp);

	/* セル名に対応する数値の取得 */
	for(i = 0; i < cell_num; i++)
	{
		/* tmp の先頭アドレスからセル名を検索 */
		scan = strstr(tmp, cells_name[i]);

		/* 開始タグ<v> と 終了タグ</v> で囲まれている部分（数字）があれば */
		if((strstr(scan, START_TAG) && strstr(scan, END_TAG)))
		{
			/* 開始タグ<v> からのアドレスを走査用ポインタに格納（文字列かどうかの判別をするため） */
			scan = strstr(scan, START_TAG);

			/* セルに格納されているデータが文字列か数字かを判別 */
			if((strncmp((scan - MOVE_TO_STRING_EXPRESSION), STRING_EXPRESSION, strlen(STRING_EXPRESSION))))
			{
				/* 数字を数値に変換(double型)して引数で渡された配列に格納 */
				intp[i] = atoi(scan + strlen(START_TAG));
				/* printf("number[%d] = %lf\n", i, doublep[i]); */
			}
			else
			{
				/* 指定したセルに数字がない場合、メッセージを表示 */
				printf("Number could not be found : Cell \"%s\"\n", cells_name[i]);
			}
		}
		else
		{
			remove_expanded_data(zip_path);
			free(cells_name);
			printf("Numbers could not be found.\n");
			exit(1);
		}

	}

	/* 作成されたZIPファイル、展開されたファイル・フォルダの削除 */
	remove_expanded_data(zip_path);

	/* セル名用に確保したメモリ領域の解放 */
	free(cells_name);
}


/************************************************************
* 関数名：xlsx2zip
* 機能：指定したファイルをコピーしてコピーファイルをZIP形式にする関数
* 引数： path XLSXファイルのパスを文字列で指定
 ************************************************************/
static char *xlsx2zip(const char *path)
{
	unsigned int insert;
	char *p;
	char copy_path[PATH_LENGTH]; /* コピーファイルのパス格納用 */
	static char zip_path[PATH_LENGTH]; /* ZIPファイルのパス格納用 */
	char copy_command[COMMAND_LENGTH]; /* OSに渡すための copy コマンド格納用 */

	/* コピー元ファイル名の拡張子の手前に _copy を挿入してコピーファイルの名前を文字列として copy_path に格納 */
	insert = strlen(path) - strlen(strchr(path, '.'));
	strcpy(copy_path, path);
	memmove(&copy_path[insert + strlen(COPY_FILENAME_INS_STR)], &copy_path[insert], strlen(&copy_path[insert - 1]));
	memcpy(&copy_path[insert], COPY_FILENAME_INS_STR, strlen(COPY_FILENAME_INS_STR));

	/* ZIPファイルの名前を文字列として zip_path に格納 */
	strcpy(zip_path, copy_path);
	p = strchr(zip_path, '.');
	p[1] = '\0';
	strcat(zip_path, "zip");

	/* OSに渡すためのcopyコマンドを文字列として copy_command に格納 */
	sprintf(copy_command, "copy %s %s", path, copy_path);

	/* コピー元のファイル名の拡張子の前に "_copy" をつけたコピーファイルを作成*/
	system(copy_command);

	/* コピーしたファイルの拡張子をZIP形式に変更 */
	rename(copy_path, zip_path);

	p = zip_path;

	/* 作成したZIP形式ファイルのパスを返す */
	return p;
}


/************************************************************
* 関数名：unzip
* 機能： 指定したZIP形式ファイルを展開する関数
* 引数： zip_path 作成したZIP形式ファイルのパスを指定
 ************************************************************/
static void unzip(const char *zip_path)
{
	char expand_command[COMMAND_LENGTH];

	/* OSに渡すためのZIPファイル展開コマンドを文字列として expand_command に格納 */
	sprintf(expand_command, "CScript.exe unzip.vbs %s", zip_path);

	/* ZIPファイルを展開 */
	system(expand_command);
}


/************************************************************
* 関数名：sheet_file_path_get
* 機能：指定されたシート名に対応するファイルのパスを返す関数
* 引数： sheet XLSXファイルのシート名を文字列で指定
*		zip_path 作成したZIP形式ファイルのパスを指定
 ************************************************************/
static char *sheet_file_path_get(char *sheet, char *zip_path)
{
	FILE *fp;
	char tmp[5000];
	char *idp;
	int id;
	static char sheet_path[256];
	char *ret;

	/* シートの ID が記述されたファイルを開く(workbook.xml) */
	if((fp = fopen("xl\\workbook.xml", "r")) == NULL)
	{
		remove_expanded_data(zip_path);
		printf("workbook.xml Not found.\n");
		exit(1);
	}
	else
	{
		while((fgets(tmp, 5000, fp)) != NULL);

		/* tmp からシート名を検索 */
		if(strstr(tmp, sheet) != NULL)
		{
			/* シート名に対応するIDを取得 */
			idp = (strstr(tmp, sheet) + strlen(sheet) + MOVE_TO_ID);
			id = atoi(idp);

			/* sheet_path にシート名に対応するファイルのパスを格納 */
			/* シート名に対応する各ファイルは、Excelのシートタブ左側から、sheet1.xml, sheet2.xml, sheet3.xml ... となっている */
			sprintf(sheet_path, "xl\\worksheets\\sheet%d.xml", id);
		}
		else
		{
			printf("Sheet Not Found.\n");
			exit(1);
		}
	}
	fclose(fp);

/*
	printf("\n\n--------------Sheet_Data-------------\n");
	printf("id = %d\n", id);
	printf("sheet_path = %s\n", sheet_path);
	printf("-------------------------------------\n");
*/

	ret = sheet_path;
	return ret;
}


/************************************************************
* 関数名：remove_expanded_data
* 機能：作成されたZIPファイルおよび展開されたファイル・フォルダを削除する関数
* 引数： zip_path 作成されたZIP形式ファイルのパスを指定
 ************************************************************/
static void remove_expanded_data(char *zip_path)
{
	char files_remove_command[COMMAND_LENGTH];

	sprintf(files_remove_command, "del /Q [Content_Types].xml %s", zip_path);

	/* ファイルの削除 */
	system(files_remove_command);

	/* ディレクトリの削除 */
	system("rd /s /q _rels docProps xl");
}
