import os
import openpyxl
import tkinter
wb = openpyxl.load_workbook("netflix_titles.xlsx")
sheet = wb["netflixtitles"]
class Movie:
    def __init__(self,type,title,director,releaseYear):
        self.type = type
        self.title = title
        self.director = director
        self.releaseYear = releaseYear
    
def createMovie(row):#Creates Movie and moves it to Class
    mov = Movie(row[0].value,row[1].value,row[2].value,row[3].value)
    return mov

def get_data_from_excel():#Pulls Data from Excel
    movies = []
    for row in sheet["A2:d25"]:
        movies.append(createMovie(row))
    return movies

def copyExcel(movies,file):
    for row in movies:
        file.writelines("\n"+ row.title)
    file.close()

def delete_from_file(fileName):
    fileName = open(fn +".txt","r",encoding="utf8")
    lines = fileName.readlines()
    answer = input("Enter Name of Movie: ")
    for i in range(len(lines)):
        if lines[i] == answer:
            del lines[i]
            new_file = open(fn+".txt","w+",encoding="utf8")
            for line in lines:
                new_file.write(line)
            new_file.close()
            break
    fileName.close()
def printMovies():
    os.system(fileName)
    file.close()
def addNew(file):
    newtext = input("Name of Movie: ")
    fileName = open(fn+".txt","a")
    fileName.writelines("\n"+ newtext)
    fileName.close()
#------------------------------------------------------------------------------------#
movies = get_data_from_excel()
answer = 0
fn = input("Enter File Name: ")
file = open(fn+".txt",'w',encoding="utf-8")
fileName = fn+".txt"
copyExcel(movies,file)
print("file created...")
while answer !=4:
    answer = int(input("Choose an Option \n1.Show Movies\n2.Add Movie\n3.Delete Movie: \n"))
    if answer ==1:
        file.close()
        printMovies()
    elif answer == 2:
        addNew(file)
    elif answer == 3:
        delete_from_file(file)
    else:
        break