package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

func indexOf(val string,arr [8]string)int{ //возвращает индекс элемента в массиве
	for i,v:=range arr{
		if val==v{
			return i;
		}
	}
	return -1;
}

func main(){
	f, _ := excelize.OpenFile("tanya.xlsx") //открываем таблицу
	var books [8] string; //названия книг
	var values [8] int64; //кол-во книг
	for i:=0; i<8; i++{
		cell:=fmt.Sprintf("A%d",i+3)
		books[i],_=f.GetCellValue("Лист3", cell)
		cell=fmt.Sprintf("B%d",i+3)
		tmp,_:=f.GetCellValue("Лист3",cell);
		values[i],_=strconv.ParseInt(tmp,10,64) //covert string to int64
	}
	//сейчас у нас есть список книг и их кол-во
	//давайте бегать по таблице!
	rows,_:=f.GetRows("Лист4")
	for i:=0;i<len(rows[15]);i++{
		if i>29{
			index:=indexOf(rows[15][1+2*i],books)
			if index!=-1{
				//есть книга, уменьшаем книги в коллекции
				values[index]--;
			}
			break
		}
		index:=indexOf(rows[15][1+3*i],books)
		if index!=-1{
			//есть книга, уменьшаем книги в коллекции
			values[index]--;
			//может есть и вторая?
			index:=indexOf(rows[16][1+3*i],books)
			if index!=-1{
				values[index]--;
			}
		}
	}
	for i:=0; i<len(values); i++{
		cell:=fmt.Sprintf("B%d",i+3)
		f.SetCellValue("Лист4",cell,values[i])
	}
	f.SaveAs("tanya1.xlsx")
}