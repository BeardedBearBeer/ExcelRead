#ifndef _EXCEL_READ_H_
#define _EXCEL_READ_H_

/* サポート関数 */
void xlsx_nums_read(void *buf, const char *path, char *sheet);
void xlsx_nums_read_from_cells(void *buf, const char *path, char *sheet, int cell_num, ...);

#endif
