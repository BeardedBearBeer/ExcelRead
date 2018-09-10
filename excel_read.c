#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stdarg.h>
#include "excel_read.h"

#define PATH_LENGTH 256
#define COMMAND_LENGTH 256
#define COPY_FILENAME_INS_STR "_copy"
#define START_TAG "<v>" /* XML�t�@�C���̐����̊J�n�^�O */
#define END_TAG "</v>" /* XML�t�@�C���̐����̏I���^�O */
#define STRING_EXPRESSION "t=\"s\""
#define MOVE_TO_ID 11
#define MOVE_TO_STRING_EXPRESSION 6

/*�^���ʗp�̍\����
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

/* �v���g�^�C�v�錾--------------------------------------------------------------------------------------------------------------------------------------------------- */
static char *xlsx2zip(const char *path);
static void unzip(const char *zip_path);
static char *sheet_file_path_get(char *sheet, char *zip_path);
static void remove_expanded_data(char *zip_path);


/************************************************************
* �֐����Fxlsx_nums_read
* �@�\�F�w�肵��XLSX�t�@�C���̃V�[�g����w�肳�ꂽ�ϐ��ɒl���i�[����֐�
* �����F buf �l���i�[�������ϐ��̃A�h���X���w��
*		path XLSX�t�@�C���̃p�X�𕶎���Ŏw��
*		sheet XLSX�t�@�C���̃V�[�g���𕶎���Ŏw��
 ************************************************************/
void xlsx_nums_read(void *buf, const char *path, char *sheet)
{
	FILE *fp;
	char *zip_path;
	int *intp = (int *)buf;
	/* double *doublep = (double *)buf; */ /* double�^������ */
	char tmp[5000]; /* XML�t�@�C���̃f�[�^�i�[�p */
	char *scan; /* �f�[�^�����p�|�C���^ */
	int i = 0;

	if(!(strstr(path, ".xlsx")))
	{
		printf("Please specify XLSX format file as path.\n");
		exit(0);
	}

	zip_path = xlsx2zip(path);
	unzip(zip_path);

	/* �w�肵���V�[�g���ɑΉ�����t�@�C���̃p�X���擾 */
	sheet = sheet_file_path_get(sheet, zip_path);

	/* �擾�����V�[�g�t�@�C�����J���đS�Ẵf�[�^�� tmp �Ɋi�[ */
	if((fp = fopen(sheet, "r")) == NULL)
	{
		printf("File Not found.\n");
		exit(1);
	}
	else
	{
		while((fgets(tmp, 5000, fp)) != NULL);

		/* �f�[�^�����p�|�C���^�� �J�n�^�O<v> ����̃A�h���X��n�� */
		scan = strstr(tmp, START_TAG);
	}
	fclose(fp);

	do {
		/* �J�n�^�O<v> �� �I���^�O</v> �ň͂܂�Ă��镔���i�����j������� */
		if((strstr(scan, START_TAG) && strstr(scan, END_TAG)))
		{
			/* �Z���Ɋi�[����Ă���f�[�^�������񂩐������𔻕� */
			if((strncmp((scan - MOVE_TO_STRING_EXPRESSION), STRING_EXPRESSION, strlen(STRING_EXPRESSION))))
			{
				/* �����𐔒l�ɕϊ�(double�^)���Ĉ����œn���ꂽ�z��Ɋi�[ */
				intp[i++] = atoi(scan + strlen(START_TAG));
				/* printf("number[%d] = %lf\n", i, doublep[i]); */
			}

			/* �f�[�^�����p�|�C���^�����̊J�n�^�O�֐i�߂� */
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

	/* �쐬���ꂽZIP�t�@�C���A�W�J���ꂽ�t�@�C���E�t�H���_�̍폜 */
	remove_expanded_data(zip_path);
}


/************************************************************
* �֐����Fxlsx_nums_read_from_cells
* �@�\�F�w�肵��XLSX�t�@�C���̃V�[�g�y�уZ������w�肳�ꂽ�ϐ��ɒl���i�[����֐�
* �����F buf �l���i�[�������ϐ��̃A�h���X���w��
*		path XLSX�t�@�C���̃p�X�𕶎���Ŏw��
*		sheet XLSX�t�@�C���̃V�[�g���𕶎���Ŏw��
*		cell_num �ΏۂƂ���Z���̌����w��
*		... �ΏۂƂ���Z�������J���}��؂�Ŏw��
 ************************************************************/
void xlsx_nums_read_from_cells(void *buf, const char *path, char *sheet, int cell_num, ...)
{
	FILE *fp;
	char *zip_path;
	int *intp = (int *)buf;
	/* double *doublep = (double *)buf; */ /* double�^������ */
	char tmp[5000]; /* XML�t�@�C���̃f�[�^�i�[�p */
	char *scan; /* �f�[�^�����p�|�C���^ */
	int i;
	va_list args;
	char **cells_name;

	if(!(strstr(path, ".xlsx")))
	{
		printf("Please specify XLSX format file as path.\n");
		exit(0);
	}

	/* �ϒ������ɓn���ꂽ�Z�����̌������������m�� */
	cells_name = (char **)malloc(cell_num * sizeof(char *));
	memset(cells_name, NULL, cell_num * sizeof(cells_name));

	/* �ϒ������ɓn���ꂽ�Z�������擾 */
	va_start(args, cell_num);
	for(i = 0; i < cell_num; i++)
	{
		cells_name[i] = va_arg(args, char *);
		/* printf("%s\n", cells_name[i]); */
	}

	/* path �Ɏw�肳�ꂽXLSX�t�@�C���̃R�s�[�t�@�C����ZIP�� */
	zip_path = xlsx2zip(path);
	/* ZIP�������t�@�C����W�J */
	unzip(zip_path);

	/* �w�肵���V�[�g���ɑΉ�����f�[�^���L�q����Ă���t�@�C���̃p�X���擾 */
	sheet = sheet_file_path_get(sheet, zip_path);

	/* �擾�����V�[�g�t�@�C�����J���đS�Ẵf�[�^�� tmp �Ɋi�[ */
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

	/* �Z�����ɑΉ����鐔�l�̎擾 */
	for(i = 0; i < cell_num; i++)
	{
		/* tmp �̐擪�A�h���X����Z���������� */
		scan = strstr(tmp, cells_name[i]);

		/* �J�n�^�O<v> �� �I���^�O</v> �ň͂܂�Ă��镔���i�����j������� */
		if((strstr(scan, START_TAG) && strstr(scan, END_TAG)))
		{
			/* �J�n�^�O<v> ����̃A�h���X�𑖍��p�|�C���^�Ɋi�[�i�����񂩂ǂ����̔��ʂ����邽�߁j */
			scan = strstr(scan, START_TAG);

			/* �Z���Ɋi�[����Ă���f�[�^�������񂩐������𔻕� */
			if((strncmp((scan - MOVE_TO_STRING_EXPRESSION), STRING_EXPRESSION, strlen(STRING_EXPRESSION))))
			{
				/* �����𐔒l�ɕϊ�(double�^)���Ĉ����œn���ꂽ�z��Ɋi�[ */
				intp[i] = atoi(scan + strlen(START_TAG));
				/* printf("number[%d] = %lf\n", i, doublep[i]); */
			}
			else
			{
				/* �w�肵���Z���ɐ������Ȃ��ꍇ�A���b�Z�[�W��\�� */
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

	/* �쐬���ꂽZIP�t�@�C���A�W�J���ꂽ�t�@�C���E�t�H���_�̍폜 */
	remove_expanded_data(zip_path);

	/* �Z�����p�Ɋm�ۂ����������̈�̉�� */
	free(cells_name);
}


/************************************************************
* �֐����Fxlsx2zip
* �@�\�F�w�肵���t�@�C�����R�s�[���ăR�s�[�t�@�C����ZIP�`���ɂ���֐�
* �����F path XLSX�t�@�C���̃p�X�𕶎���Ŏw��
 ************************************************************/
static char *xlsx2zip(const char *path)
{
	unsigned int insert;
	char *p;
	char copy_path[PATH_LENGTH]; /* �R�s�[�t�@�C���̃p�X�i�[�p */
	static char zip_path[PATH_LENGTH]; /* ZIP�t�@�C���̃p�X�i�[�p */
	char copy_command[COMMAND_LENGTH]; /* OS�ɓn�����߂� copy �R�}���h�i�[�p */

	/* �R�s�[���t�@�C�����̊g���q�̎�O�� _copy ��}�����ăR�s�[�t�@�C���̖��O�𕶎���Ƃ��� copy_path �Ɋi�[ */
	insert = strlen(path) - strlen(strchr(path, '.'));
	strcpy(copy_path, path);
	memmove(&copy_path[insert + strlen(COPY_FILENAME_INS_STR)], &copy_path[insert], strlen(&copy_path[insert - 1]));
	memcpy(&copy_path[insert], COPY_FILENAME_INS_STR, strlen(COPY_FILENAME_INS_STR));

	/* ZIP�t�@�C���̖��O�𕶎���Ƃ��� zip_path �Ɋi�[ */
	strcpy(zip_path, copy_path);
	p = strchr(zip_path, '.');
	p[1] = '\0';
	strcat(zip_path, "zip");

	/* OS�ɓn�����߂�copy�R�}���h�𕶎���Ƃ��� copy_command �Ɋi�[ */
	sprintf(copy_command, "copy %s %s", path, copy_path);

	/* �R�s�[���̃t�@�C�����̊g���q�̑O�� "_copy" �������R�s�[�t�@�C�����쐬*/
	system(copy_command);

	/* �R�s�[�����t�@�C���̊g���q��ZIP�`���ɕύX */
	rename(copy_path, zip_path);

	p = zip_path;

	/* �쐬����ZIP�`���t�@�C���̃p�X��Ԃ� */
	return p;
}


/************************************************************
* �֐����Funzip
* �@�\�F �w�肵��ZIP�`���t�@�C����W�J����֐�
* �����F zip_path �쐬����ZIP�`���t�@�C���̃p�X���w��
 ************************************************************/
static void unzip(const char *zip_path)
{
	char expand_command[COMMAND_LENGTH];

	/* OS�ɓn�����߂�ZIP�t�@�C���W�J�R�}���h�𕶎���Ƃ��� expand_command �Ɋi�[ */
	sprintf(expand_command, "CScript.exe unzip.vbs %s", zip_path);

	/* ZIP�t�@�C����W�J */
	system(expand_command);
}


/************************************************************
* �֐����Fsheet_file_path_get
* �@�\�F�w�肳�ꂽ�V�[�g���ɑΉ�����t�@�C���̃p�X��Ԃ��֐�
* �����F sheet XLSX�t�@�C���̃V�[�g���𕶎���Ŏw��
*		zip_path �쐬����ZIP�`���t�@�C���̃p�X���w��
 ************************************************************/
static char *sheet_file_path_get(char *sheet, char *zip_path)
{
	FILE *fp;
	char tmp[5000];
	char *idp;
	int id;
	static char sheet_path[256];
	char *ret;

	/* �V�[�g�� ID ���L�q���ꂽ�t�@�C�����J��(workbook.xml) */
	if((fp = fopen("xl\\workbook.xml", "r")) == NULL)
	{
		remove_expanded_data(zip_path);
		printf("workbook.xml Not found.\n");
		exit(1);
	}
	else
	{
		while((fgets(tmp, 5000, fp)) != NULL);

		/* tmp ����V�[�g�������� */
		if(strstr(tmp, sheet) != NULL)
		{
			/* �V�[�g���ɑΉ�����ID���擾 */
			idp = (strstr(tmp, sheet) + strlen(sheet) + MOVE_TO_ID);
			id = atoi(idp);

			/* sheet_path �ɃV�[�g���ɑΉ�����t�@�C���̃p�X���i�[ */
			/* �V�[�g���ɑΉ�����e�t�@�C���́AExcel�̃V�[�g�^�u��������Asheet1.xml, sheet2.xml, sheet3.xml ... �ƂȂ��Ă��� */
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
* �֐����Fremove_expanded_data
* �@�\�F�쐬���ꂽZIP�t�@�C������ѓW�J���ꂽ�t�@�C���E�t�H���_���폜����֐�
* �����F zip_path �쐬���ꂽZIP�`���t�@�C���̃p�X���w��
 ************************************************************/
static void remove_expanded_data(char *zip_path)
{
	char files_remove_command[COMMAND_LENGTH];

	sprintf(files_remove_command, "del /Q [Content_Types].xml %s", zip_path);

	/* �t�@�C���̍폜 */
	system(files_remove_command);

	/* �f�B���N�g���̍폜 */
	system("rd /s /q _rels docProps xl");
}
