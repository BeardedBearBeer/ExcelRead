#ifndef _EXCEL_READ_H_
#define _EXCEL_READ_H_

/* Name :  xlsx_nums_read
    @brief 指定されたXLSX形式ファイルのシートから全ての数値データを buf に格納します
    @param buf 値の格納先
    @param path XLSX形式ファイルのパス
    @param sheet 対象とするExcelのシート名
*/
void xlsx_nums_read(void *buf, const char *path, char *sheet);

/* Name :  xlsx_nums_read_from_cells
    @brief 指定されたXLSX形式ファイルのセルから数値データを buf に格納します
    @param buf 値の格納先
    @param path XLSX形式ファイルのパス
    @param sheet 対象とするExcelのシート名
    @param cell_num 対象とするセルの個数
    @param ... 対象とするするセル名
*/
void xlsx_nums_read_from_cells(void *buf, const char *path, char *sheet, int cell_num, ...);

#endif
