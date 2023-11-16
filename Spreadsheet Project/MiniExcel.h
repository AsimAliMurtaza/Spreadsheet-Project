#pragma once
#include <iostream>
#include <windows.h>
#include <cmath>
#include <string>
#include <sstream>
#include <conio.h>
#include <fstream>
#include <cstdlib>
#include <vector>

#define INT_MIN -2147483648
#define INT_MAX 2147483647
using namespace std;

enum Color
{
    Blue,
    Red,
    White
};

class Cell
{
    string value;
    int x;
    int y;
    string valueType;
    Color color;

    Cell(int val, int xVal, int yVal, string Type)
    {
        value = val;
        x = xVal;
        y = yVal;
        color = Blue;
        valueType = Type;
    }

    Cell()
    {
        x = 0;
        y = 0;
        color = Blue;
        value = "    ";
    }

public:

    void setSelectCellColor()
    {
        color = White;
    }

    void setDeselectCellColor()
    {
        color = Blue;
    }

    void setNullValue()
    {
        value = "    ";
    }

    void setValue(string val)
    {
        string str = "";

        for (int i = 0; i < 4 && i < val.length(); i++)
        {
            str += val[i];
        }

        if (val.length() < 4)
        {
            for (int i = 0; i < 4 - val.length(); i++)
            {
                str += " ";
            }
        }
        this->value = str;
    }

    int getX()
    {
        return x;
    }

    int getY()
    {
        return y;
    }

    string getData()
    {
        return value;
    }

    int getColor()
    {
        if (color == Red)
        {
            return 4;
        }

        else if (color == Blue)
        {
            return 3;
        }
        else
        {
            return 15;
        }
    }

    friend class Node;
};

class Node
{

public:

    Node* left;
    Node* right;
    Node* up;
    Node* down;
    Cell* cell;



    Node(Cell* value)
    {
        cell = value;
        right = nullptr;
        left = nullptr;
        up = nullptr;
        down = nullptr;
    }

    void calculateNodeLocation()
    {
        Node* newNode = this;
        int count = 0;

        while (newNode->up != nullptr)
        {
            count++;
            newNode = newNode->up;
        }
        cell->y = count;
        count = 0;

        while (newNode->left != nullptr)
        {
            count++;
            newNode = newNode->left;
        }
        cell->x = count;
    }
    Node()
    {
        cell = new Cell();
        right = nullptr;
        left = nullptr;
        up = nullptr;
        down = nullptr;
    }

    friend class MiniExcel;
};

class Iterator
{

public:

    Node* i = nullptr;

    Iterator()
    {
        i = nullptr;
    }

    Iterator(Node* n)
    {
        i = n;
    }



    Iterator operator--()
    {
        if (i->up != nullptr) {
            i = i->up;
        }
        return *this;
    }
    Iterator operator++()
    {
        if (i->down != nullptr)
        {
            i = i->down;
        }
        return *this;
    }



    Iterator operator--(int)
    {
        if (i->left != nullptr) {
            i = i->left;
        }
        return *this;
    }
    Iterator operator++(int)
    {
        if (i->right != nullptr) {
            i = i->right;
        }
        return *this;
    }



    bool operator!=(Iterator n)
    {
        return (i != n.i);
    }
    bool operator==(Iterator n)
    {
        return (i == n.i);
    }
    friend class MiniExcel;
};

class MiniExcel
{
    friend class Node;

private:

    Node* selectedNode;
    vector<string> copiedData;
    char copiedLocation = ' ';

public:

    MiniExcel(int row, int col) {
        selectedNode = new Node();

        for (int i = 0; i < col; i++)
        {
            extendColumnFromRight();
        }

        for (int i = 0; i < row; i++)
        {
            extendRowFromBottom();
        }
    }

    Node* getTopRightNode()
    {
        Node* temp = selectedNode;

        while (temp->up)
        {
            temp = temp->up;
        }

        while (temp->right)
        {
            temp = temp->right;
        }
        return temp;
    }

    Node* getBottomLeftNode()
    {
        Node* temp = selectedNode;

        while (temp->left)
        {
            temp = temp->left;
        }

        while (temp->down)
        {
            temp = temp->down;
        }
        return temp;
    }

    Node* getLeftMostNode()
    {
        Node* temp = selectedNode;

        while (temp->left != nullptr)
        {
            temp = temp->left;
        }
        return temp;
    }

    Node* getTopMostNode()
    {
        Node* temp = selectedNode;

        while (temp->up != nullptr)
        {
            temp = temp->up;
        }
        return temp;
    }

    Node* GetNodeAtTopLeft()
    {
        Node* temp = selectedNode;

        while (temp->left)
        {
            temp = temp->left;
        }

        while (temp->up)
        {
            temp = temp->up;
        }
        return temp;
    }

    Node* getSelectedNode()
    {
        return selectedNode;
    }

    void moveRight()
    {
        if (selectedNode->right != nullptr)
        {
            selectedNode->cell->setDeselectCellColor();
            selectedNode = selectedNode->right;
            selectedNode->cell->setSelectCellColor();
            displayExcel();
        }
        else
        {
            extendColumnFromRight();
        }
    }

    void moveLeft()
    {
        if (selectedNode->left != nullptr)
        {
            selectedNode->cell->setDeselectCellColor();
            selectedNode = selectedNode->left;
            selectedNode->cell->setSelectCellColor();
            displayExcel();
        }
    }

    void moveUp()
    {
        if (selectedNode->up != nullptr)
        {
            selectedNode->cell->setDeselectCellColor();
            selectedNode = selectedNode->up;
            selectedNode->cell->setSelectCellColor();
            displayExcel();
        }
    }

    void moveDown()
    {
        if (selectedNode->down != nullptr)
        {
            selectedNode->cell->setDeselectCellColor();
            selectedNode = selectedNode->down;
            selectedNode->cell->setSelectCellColor();
            displayExcel();
        }
        else
        {
            extendRowFromBottom();
        }
    }

    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);

    void displayExcel()
    {
        system("cls");
        Node* tempRow = GetNodeAtTopLeft();

        while (tempRow != nullptr)
        {
            Node* tempCol = tempRow;
            while (tempCol != nullptr)
            {
                tempCol->calculateNodeLocation();
                SetConsoleTextAttribute(hConsole, tempCol->cell->getColor());
                gotoxy((tempCol->cell->getX() * 6), (tempCol->cell->getY() * 4));
                cout << "+----+" << endl;
                gotoxy((tempCol->cell->getX() * 6), (tempCol->cell->getY() * 4) + 1);
                cout << "|    |" << endl;
                gotoxy((tempCol->cell->getX() * 6), (tempCol->cell->getY() * 4) + 2);
                cout << "|" << tempCol->cell->getData() << "|" << endl;
                gotoxy((tempCol->cell->getX() * 6), (tempCol->cell->getY() * 4) + 3);
                cout << "|____|" << endl;
                if (tempCol->down != nullptr)
                {
                    gotoxy((tempCol->cell->getX() * 6), (tempCol->cell->getY() * 4) + 3);
                    cout << "|____|" << endl;
                }
                tempCol = tempCol->right;
            }
            tempRow = tempRow->down;
        }
    }

    char getCharAtxy(short int x, short int y)
    {
        CHAR_INFO ci;
        COORD xy = { 0, 0 };
        SMALL_RECT rect = { x, y, x, y };
        COORD coordBufSize;
        coordBufSize.X = 1;
        coordBufSize.Y = 1;
        return ReadConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), &ci, coordBufSize, xy, &rect) ? ci.Char.AsciiChar : ' ';
    }

    void gotoxy(int x, int y)
    {
        COORD coordinates;
        coordinates.X = x;
        coordinates.Y = y;
        SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
    }

    // baki functions
    void insertData()
    {
        string x = "";
        cout << "Enter value: ";
        cin >> x;
        selectedNode->cell->setValue(x);
        displayExcel();
    }

    void extendRowFromBottom()
    {
        Node* temp = getBottomLeftNode();

        while (temp)
        {
            Node* newNode = new Node();
            temp->down = newNode;
            temp->down->up = temp;
            temp = temp->right;
        }

        temp = getBottomLeftNode();

        while (temp->up->right)
        {
            temp->right = temp->up->right->down;
            temp->right->left = temp;
            temp = temp->right;
        }
    }

    void extendColumnFromRight()
    {
        Node* temp = getTopRightNode();

        while (temp)
        {
            Node* newNode = new Node();
            temp->right = newNode;
            temp->right->left = temp;
            temp = temp->down;
        }

        temp = getTopRightNode();

        while (temp->left->down)
        {
            temp->down = temp->left->down->right;
            temp->down->up = temp;
            temp = temp->down;
        }
    }

    void clearRow()
    {
        Node* newNode = getLeftMostNode();

        while (newNode->right != nullptr)
        {
            newNode->cell->setNullValue();
            newNode = newNode->right;
        }
        displayExcel();
    }

    void clearColumn()
    {
        Node* newNode = getTopMostNode();

        while (newNode->down != nullptr)
        {
            newNode->cell->setNullValue();
            newNode = newNode->down;
        }

        displayExcel();
    }

    void deleteRow()
    {
        Node* temp = getLeftMostNode();

        selectedNode = selectedNode->up;
        Node* nextTemp;
        while (temp != nullptr)
        {
            if (temp->up == nullptr)
            {
                temp->down->up = nullptr;
            }
            else if (temp->down == nullptr)
            {
                temp->up->down = nullptr;
            }
            else
            {
                temp->up->down = temp->down;
                temp->down->up = temp->up;
            }
            nextTemp = temp->right;
            delete temp;
            temp = nextTemp;
        }
        displayExcel();
    }

    void deleteColumn()
    {
        Node* temp = getTopMostNode();

        while (temp->down != nullptr)
        {
            if (temp->left == nullptr)
            {
                temp->right->left = nullptr;
            }
            else if (temp->right == nullptr)
            {
                temp->left->right = nullptr;
            }
            else
            {
                temp->right->left = temp->left;
                temp->left->right = temp->right;
            }
            temp = temp->down;
        }
        displayExcel();
    }

    void insertColumnAtRight()
    {
        Node* temp = getTopMostNode();
        Node* nextTemp;

        while (temp)
        {
            if (temp->right)
            {
                Node* newNode = new Node();
                nextTemp = temp->right;
                temp->right = newNode;
                temp->right->left = temp;
                temp->right->right = nextTemp;
                temp->right->right->left = temp->right;
            }
            else
            {
                Node* newNode = new Node();
                temp->right = newNode;
                temp->right->left = temp;
                temp->right->right = nullptr;
            }
            temp = temp->down;
        }

        temp = getTopMostNode();

        while (temp->down)
        {
            temp->right->down = temp->down->right;
            temp->down->right->up = temp->right;
            temp = temp->down;
        }
        displayExcel();
    }

    void insertColumnAtLeft()
    {
        Node* temp = getTopMostNode();
        Node* nextTemp;

        while (temp)
        {
            if (temp->left)
            {
                Node* newNode = new Node();
                nextTemp = temp->left;
                temp->left = newNode;
                temp->left->right = temp;
                temp->left->left = nextTemp;
                temp->left->left->right = temp->left;
            }
            else
            {
                Node* newNode = new Node();
                temp->left = newNode;
                temp->left->right = temp;
                temp->left->left = nullptr;
            }
            temp = temp->down;
        }

        temp = getTopMostNode();

        while (temp->down)
        {
            temp->left->down = temp->down->left;
            temp->down->left->up = temp->left;
            temp = temp->down;
        }
        displayExcel();
    }

    void insertRowAbove()
    {
        Node* temp = getLeftMostNode();
        Node* nextTemp;

        while (temp)
        {
            if (temp->up)
            {
                Node* newNode = new Node();
                nextTemp = temp->up;
                temp->up = newNode;
                newNode->down = temp;
                newNode->up = nextTemp;
                nextTemp->down = newNode;
            }
            else
            {
                Node* newNode = new Node();
                temp->up = newNode;
                newNode->down = temp;
                newNode->up = nullptr;
            }
            temp = temp->right;
        }

        temp = getLeftMostNode();

        while (temp->right)
        {
            temp->up->right = temp->right->up;
            temp->right->up->left = temp->up;
            temp = temp->right;
        }
        displayExcel();
    }

    void insertRowBelow()
    {
        Node* temp = getLeftMostNode();
        Node* nextTemp;

        while (temp)
        {
            if (temp->down)
            {
                Node* newNode = new Node();
                nextTemp = temp->down;
                temp->down = newNode;
                newNode->up = temp;
                newNode->down = nextTemp;
                nextTemp->up = newNode;
            }
            else
            {
                Node* newNode = new Node();
                temp->down = newNode;
                newNode->up = temp;
                newNode->down = nullptr;
            }
            temp = temp->right;
        }

        temp = getLeftMostNode();

        while (temp->right)
        {
            temp->down->right = temp->right->down;
            temp->right->down->left = temp->down;
            temp = temp->right;
        }
        displayExcel();
    }

    void insertCellByRightShift()
    {
        Node* temp = selectedNode;
        Node* lastNode = temp;
        string valueOfnextCell = "    ";
        string valueOfcurrentCell = "    ";

        while (lastNode->right)
        {
            lastNode = lastNode->right;
        }

        if (lastNode->cell->getData() != "    ")
        {
            extendColumnFromRight();
        }

        valueOfcurrentCell = temp->cell->getData();

        while (temp->right)
        {
            valueOfnextCell = temp->right->cell->getData();
            temp->right->cell->setValue(valueOfcurrentCell);
            valueOfcurrentCell = valueOfnextCell;
            temp = temp->right;
        }
        temp = selectedNode;
        temp->cell->setNullValue();
        displayExcel();
    }

    void insertCellByDownShift()
    {
        Node* temp = selectedNode;
        Node* lastNode = temp;
        string valueOfnextCell = "    ";
        string valueOfcurrentCell = "    ";

        while (lastNode->down)
        {
            lastNode = lastNode->down;
        }

        if (lastNode->cell->getData() != "    ")
        {
            extendRowFromBottom();
        }

        valueOfcurrentCell = temp->cell->getData();

        while (temp->down)
        {
            valueOfnextCell = temp->down->cell->getData();
            temp->down->cell->setValue(valueOfcurrentCell);
            valueOfcurrentCell = valueOfnextCell;
            temp = temp->down;
        }
        temp = selectedNode;
        temp->cell->setNullValue();
        displayExcel();
    }

    void deleteCellByLeftShift()
    {
        Node* temp = selectedNode;
        Node* lastNode = temp;
        string valueOfnextCell = "    ";
        string valueOfcurrentCell = "    ";

        while (lastNode->right)
        {
            lastNode = lastNode->right;
        }

        valueOfcurrentCell = lastNode->cell->getData();

        while (lastNode != temp)
        {
            valueOfnextCell = lastNode->left->cell->getData();
            lastNode->left->cell->setValue(valueOfcurrentCell);
            valueOfcurrentCell = valueOfnextCell;
            lastNode = lastNode->left;
        }

        while (lastNode->right)
        {
            lastNode = lastNode->right;
        }
        lastNode->cell->setNullValue();
        displayExcel();
    }

    void deleteCellByDownShift()
    {
        Node* temp = selectedNode;
        Node* lastNode = temp;
        string valueOfnextCell = "    ";
        string valueOfcurrentCell = "    ";

        while (lastNode->down)
        {
            lastNode = lastNode->down;
        }

        valueOfcurrentCell = lastNode->cell->getData();

        while (lastNode != temp)
        {
            valueOfnextCell = lastNode->up->cell->getData();
            lastNode->up->cell->setValue(valueOfcurrentCell);
            valueOfcurrentCell = valueOfnextCell;
            lastNode = lastNode->up;
        }

        while (lastNode->down)
        {
            lastNode = lastNode->down;
        }
        lastNode->cell->setNullValue();
        displayExcel();
    }

    void swapTwoCells()
    {
        Node* temp = selectedNode;

        if (temp->right)
        {
            string tempData = temp->right->cell->getData();
            temp->right->cell->setValue(temp->cell->getData());
            temp->cell->setValue(tempData);
        }
        else if (temp->right == nullptr)
        {
            insertColumnAtRight();
            string tempData = temp->right->cell->getData();
            temp->right->cell->setValue(temp->cell->getData());
            temp->cell->setValue(tempData);
        }
        displayExcel();
    }

    void copy(Node* start, Node* end)
    {
        if (start->cell->getX() == end->cell->getX())
        {
            while (start->up != end)
            {
                copiedData.push_back(start->cell->getData());
                start = start->down;
            }
            copiedLocation = 'c';
        }
        else if (start->cell->getY() == end->cell->getY())
        {
            while (start->left != end)
            {
                copiedData.push_back(start->cell->getData());
                start = start->right;
            }
            copiedLocation = 'r';
        }
    }

    void cut(Node* start, Node* end)
    {
        if (start->cell->getX() == end->cell->getX())
        {
            while (start->up != end)
            {
                copiedData.push_back(start->cell->getData());
                start->cell->setNullValue();
                start = start->down;
            }
            copiedLocation = 'c';
        }
        if (start->cell->getY() == end->cell->getY())
        {
            while (start->left != end)
            {
                copiedData.push_back(start->cell->getData());
                start->cell->setNullValue();
                start = start->right;
            }
            copiedLocation = 'r';
        }
        displayExcel();
    }

    void paste()
    {
        Node* temp = selectedNode;
        Node* extraTemp = temp;
        int count = 0;
        if (copiedLocation == 'r')
        {
            while (extraTemp->right)
            {
                count++;
                extraTemp = extraTemp->right;
            }
            if (count < copiedData.size())
            {
                for (int i = 0; i < copiedData.size() - count - 1; i++)
                {
                    extendColumnFromRight();
                }
            }
            count = 0;
            while (count < copiedData.size())
            {
                temp->cell->setValue(copiedData[count]);
                count++;
                temp = temp->right;
            }
            copiedData.clear();
        }
        else if (copiedLocation == 'c')
        {
            while (extraTemp->down)
            {
                count++;
                extraTemp = extraTemp->down;
            }
            if (count < copiedData.size())
            {
                for (int i = 0; i < copiedData.size() - count - 1; i++)
                {
                    extendRowFromBottom();
                }
            }
            count = 0;
            while (count < copiedData.size())
            {
                temp->cell->setValue(copiedData[count]);
                count++;
                temp = temp->down;
            }
            copiedData.clear();
        }
        displayExcel();
    }

    void sum(Node* start, Node* end)
    {
        int sum = 0;

        if (start->cell->getX() == end->cell->getX())
        {
            while (start->up != end)
            {
                copiedData.push_back(start->cell->getData());
                start = start->down;
            }
            copiedLocation = 'c';
        }
        else if (start->cell->getY() == end->cell->getY())
        {
            while (start->left != end)
            {
                copiedData.push_back(start->cell->getData());
                start = start->right;
            }
            copiedLocation = 'r';
        }

        for (const string& str : copiedData)
        {
            istringstream iss(str);
            int num = 0;

            if (iss >> num)
            {
                sum += num;
            }
        }
        copiedData.clear();
        selectedNode->cell->setValue(to_string(sum));
        displayExcel();
    }

    void average(Node* start, Node* end)
    {
        int sum = 0;
        int count = 0;

        if (start->cell->getX() == end->cell->getX())
        {
            while (start->up != end)
            {
                copiedData.push_back(start->cell->getData());
                start = start->down;
            }
            copiedLocation = 'c';
        }
        else if (start->cell->getY() == end->cell->getY())
        {
            while (start->left != end)
            {
                copiedData.push_back(start->cell->getData());
                start = start->right;
            }
            copiedLocation = 'r';
        }

        for (const string& str : copiedData)
        {
            istringstream iss(str);
            int num = 0;

            if (iss >> num)
            {
                count++;
                sum += num;
            }
        }
        selectedNode->cell->setValue(to_string(sum / count));
        displayExcel();
    }

    void count(Node* start, Node* end)
    {
        if (start->cell->getX() == end->cell->getX())
        {
            while (start->up != end)
            {
                if (start->cell->getData() != "    ")
                {
                    copiedData.push_back(start->cell->getData());
                }
                start = start->down;
            }
            copiedLocation = 'c';
        }
        else if (start->cell->getY() == end->cell->getY())
        {
            while (start->left != end)
            {
                if (start->cell->getData() != "    ")
                {
                    copiedData.push_back(start->cell->getData());
                }
                start = start->right;
            }
            copiedLocation = 'r';
        }
        selectedNode->cell->setValue(to_string(copiedData.size()));
        copiedData.clear();
        displayExcel();
    }

    void minimum(Node* start, Node* end)
    {
        if (start->cell->getX() == end->cell->getX())
        {
            while (start->up != end)
            {
                if (start->cell->getData() != "    ")
                {
                    copiedData.push_back(start->cell->getData());
                }
                start = start->down;
            }
            copiedLocation = 'c';
        }
        else if (start->cell->getY() == end->cell->getY())
        {
            while (start->left != end)
            {
                if (start->cell->getData() != "    ")
                {
                    copiedData.push_back(start->cell->getData());
                }
                start = start->right;
            }
            copiedLocation = 'r';
        }

        int minVal = INT_MAX;

        for (const string& str : copiedData)
        {
            istringstream iss(str);
            int num = 0;
            if (iss >> num)
            {
                minVal = min(minVal, num);
            }
        }
        selectedNode->cell->setValue(to_string(minVal));
        copiedData.clear();
        displayExcel();
    }

    void maximum(Node* start, Node* end)
    {
        if (start->cell->getX() == end->cell->getX())
        {
            while (start->up != end)
            {
                if (start->cell->getData() != "    ")
                {
                    copiedData.push_back(start->cell->getData());
                }
                start = start->down;
            }
            copiedLocation = 'c';
        }
        else if (start->cell->getY() == end->cell->getY())
        {
            while (start->left != end)
            {
                if (start->cell->getData() != "    ")
                {
                    copiedData.push_back(start->cell->getData());
                }
                start = start->right;
            }
            copiedLocation = 'r';
        }

        int maxVal = INT_MIN;

        for (const string& str : copiedData)
        {
            istringstream iss(str);
            int num = 0;
            if (iss >> num)
            {
                maxVal = max(maxVal, num);
            }
        }
        selectedNode->cell->setValue(to_string(maxVal));
        copiedData.clear();
        displayExcel();
    }

    Node* getNodeByXY(int x, int y)
    {
        Node* tempRow = GetNodeAtTopLeft();

        for (int i = 0; i < x; i++)
        {
            tempRow = tempRow->right;
        }

        for (int i = 0; i < y; i++)
        {
            tempRow = tempRow->down;
        }

        return tempRow;

        while (tempRow)
        {
            Node* tempCol = tempRow;
            while (tempCol)
            {
                if (tempCol->cell->getX() == x && tempCol->cell->getY() == y)
                {
                    return tempCol;
                }
                tempCol = tempCol->right;
            }
            tempRow = tempRow->down;
        }
    }
};

void saveToFile(string file, MiniExcel excel)
{
    ofstream out(file);

    if (!out.is_open())
    {
        return;
    }
    Node* tempRow = excel.GetNodeAtTopLeft();

    while (tempRow)
    {
        Node* tempCol = tempRow;

        while (tempCol)
        {
            out << tempCol->cell->getData() << ",";
            tempCol = tempCol->right;
        }
        out << "\n";
        tempRow = tempRow->down;
    }
    out.close();
}

void saveRowsAndColumn(string file, MiniExcel excel)
{
    ofstream out(file);

    if (!out.is_open())
    {
        return;
    }

    Node* temp = excel.GetNodeAtTopLeft();

    int rowCount = 0;
    int colCount = 0;

    while (temp)
    {
        rowCount++;
        temp = temp->down;
    }

    temp = excel.GetNodeAtTopLeft();

    while (temp)
    {
        colCount++;
        temp = temp->right;
    }
    out << rowCount << "," << colCount;
    out.close();
}

string parseItems(string line, int count)
{
    int commaCount = 1;
    string item;
    for (int x = 0; x < line.length(); x++)
    {
        if (line[x] == ',')
        {
            commaCount = commaCount + 1;
        }
        else if (commaCount == count)
        {
            item = item + line[x];
        }
    }
    return item;
}

vector<int> loadData()
{
    vector<int> rowsNCols;
    fstream file;
    string word;
    file.open("rowSheet.txt", ios::in);

    while (!file.eof())
    {
        getline(file, word);
        if (word == "")
        {
            break;
        }
        rowsNCols.push_back(stoi(parseItems(word, 1)));
        rowsNCols.push_back(stoi(parseItems(word, 2)));
    }
    file.close();
    return rowsNCols;
}

void loadFromFile(string file, MiniExcel excel)
{
    int row = 0;
    ifstream in(file);

    if (!in.is_open())
    {
        return;
    }

    string line;

    while (getline(in, line))
    {
        istringstream iss(line);
        string token;
        int col = 0;

        while (getline(iss, token, ','))
        {
            Node* currentCell = excel.getNodeByXY(col, row);

            if (currentCell != nullptr)
            {
                currentCell->cell->setValue(token);
            }
            col++;
        }
        row++;
    }
    in.close();
}