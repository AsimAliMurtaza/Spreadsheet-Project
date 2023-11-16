#include <iostream>
#include "MiniExcel.h"

int rows = 0;
int cols = 0;

int main()
{
	Node* start = new Node();
	Node* end = new Node();
	vector<int> data = loadData();

	rows = data[0];
	cols = data[1];

	MiniExcel excel(rows, cols);
	loadFromFile("excelSheet.txt", excel);
	excel.displayExcel();

	while (true)
	{
		char key = _getch();

		if (key == ',')
		{
			start = excel.getSelectedNode();
		}
		if (key == '.')
		{
			end = excel.getSelectedNode();
		}
		if (key == 75)
		{
			excel.moveLeft();
		}
		if (key == 77)
		{
			excel.moveRight();
		}
		if (key == 72)
		{
			excel.moveUp();
		}
		if (key == 80)
		{
			excel.moveDown();
		}

		if (key == 32)
		{
			excel.insertData();
		}
		if (key == 'z')
		{
			excel.swapTwoCells();
		}

		if (key == 'q')
		{
			excel.clearRow();
		}
		if (key == 'w')
		{
			excel.clearColumn();
		}
		if (key == 'e')
		{
			excel.deleteRow();
		}
		if (key == 'r')
		{
			excel.deleteColumn();
		}
		if (key == 't')
		{
			excel.insertColumnAtRight();
		}
		if (key == 'y')
		{
			excel.insertColumnAtLeft();
		}
		if (key == 'u')
		{
			excel.insertRowAbove();
		}
		if (key == 'i')
		{
			excel.insertRowBelow();
		}
		if (key == 'a')
		{
			excel.insertCellByRightShift();
		}
		if (key == 's')
		{
			excel.insertCellByDownShift();
		}
		if (key == 'd')
		{
			excel.deleteCellByLeftShift();
		}
		if (key == 'f')
		{
			excel.deleteCellByDownShift();
		}

		if (key == 'c')
		{
			excel.copy(start, end);
		}
		if (key == 'x')
		{

			excel.cut(start, end);
		}
		if (key == 'v')
		{
			excel.paste();
		}

		if (key == '1')
		{
			excel.sum(start, end);
		}
		if (key == '2')
		{
			excel.average(start, end);
		}
		if (key == '3')
		{
			excel.count(start, end);
		}
		if (key == '4')
		{
			excel.minimum(start, end);
		}
		if (key == '5')
		{
			excel.maximum(start, end);
		}
		saveToFile("excelSheet.txt", excel);
		saveRowsAndColumn("rowSheet.txt", excel);
	}
}