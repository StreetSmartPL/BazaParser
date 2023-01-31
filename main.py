import openpyxl
import json
import os
from dataclasses import dataclass
from dataclasses_json import dataclass_json


@dataclass_json
@dataclass
class LangPyt:
  Pytanie: str # Pytanie
  OdpowiedzA: str # Odpowiedź A
  OdpowiedzB: str # Odpowiedź B
  OdpowiedzC: str # Odpowiedź C
  
@dataclass_json
@dataclass
class Kategorie:
  Kat_A: bool
  Kat_B: bool
  Kat_C: bool
  Kat_D: bool
  Kat_T: bool
  Kat_AM: bool
  Kat_A1: bool
  Kat_A2: bool
  Kat_B1: bool
  Kat_C1: bool
  Kat_D1: bool

@dataclass_json
@dataclass
class Pytanie:
  NazwaPytania: str # Nazwa pytania
  NumerPytania: str # Numer pytania
  PytaniePL: LangPyt # Pytanie
  PytanieENG: LangPyt # Pytanie ENG
  PytanieDE: LangPyt # Pytanie DE
  PoprawnaOdp: str # Poprawna odp
  Media: str # Media
  ZakresStruktury: str # Zakres struktury
  LiczbaPunktow: str # Liczba punktów
  Kategorie: Kategorie # Kategorie
  NazwaBloku: str # Nazwa bloku
  ZrodloPytania: str # Źródło pytania
  OCoChcemyZapytac: str # O co chcemy zapytać
  JakiMaZwiazekZBezpieczenstwem: str # Jaki ma związek z bezpieczeństwem
  Status: str # Status
  Podmiot: str # Podmiot


def removeNotFuckingImportantFiles(arrayofnames):
  for name in arrayofnames:
    try:
      os.remove("./videos/" + name)
      print("% s removed successfully" % name)
    except OSError as error:
      print(error)
      print("File path can not be removed")

def read_excel():
  temp = {}
  wb = openpyxl.load_workbook('Baza.xlsx')
  sheet = wb['Treść pytania']
  tmp = []
  for row in sheet.iter_rows(min_row=2, max_col=33):
    if row[0].value is None:
      continue
    pytanie = Pytanie(
      NazwaPytania=row[0].value,
      NumerPytania=row[1].value,
      PytaniePL=LangPyt( 
        Pytanie=row[2].value,
        OdpowiedzA=row[3].value,
        OdpowiedzB=row[4].value,
        OdpowiedzC=row[5].value
      ),
      PytanieENG=LangPyt( 
        Pytanie=row[6].value,
        OdpowiedzA=row[7].value,
        OdpowiedzB=row[8].value,
        OdpowiedzC=row[9].value
      ),
      PytanieDE=LangPyt( 
        Pytanie=row[10].value,
        OdpowiedzA=row[11].value,
        OdpowiedzB=row[12].value,
        OdpowiedzC=row[13].value
      ),
      PoprawnaOdp=row[14].value,
      Media=row[15].value,
      ZakresStruktury=row[16].value,
      LiczbaPunktow=row[17].value,
      Kategorie=Kategorie(
        Kat_A = row[18].value.__contains__("A"),
        Kat_B = row[18].value.__contains__("B"),
        Kat_C = row[18].value.__contains__("C"),
        Kat_D = row[18].value.__contains__("D"),
        Kat_T = row[18].value.__contains__("T"),
        Kat_AM = row[18].value.__contains__("AM"),
        Kat_A1 = row[18].value.__contains__("A1"),
        Kat_A2 = row[18].value.__contains__("A2"),
        Kat_B1 = row[18].value.__contains__("B1"),
        Kat_C1 = row[18].value.__contains__("C1"),
        Kat_D1 = row[18].value.__contains__("D1")
      ),
      NazwaBloku=row[19].value,
      ZrodloPytania=row[20].value,
      OCoChcemyZapytac=row[21].value,
      JakiMaZwiazekZBezpieczenstwem=row[22].value,
      Status=row[23].value,
      Podmiot=row[24].value
    )
    temp[pytanie.NazwaPytania] = pytanie
    for i in range(25, 33):
      if row[i].value != "":
        tmp.append(row[i].value)
  removeNotFuckingImportantFiles(tmp)
  return temp


def main():
  questions = read_excel()

  with open('data.json', 'w') as outfile:
    json.dump(questions, outfile, default=lambda o: o.__dict__, ensure_ascii=False)
  with open('kategorie.json','w') as outfile:
    kategorie = {}
    for key, value in questions.items():
      for k, v in value.Kategorie.__dict__.items():
        if v:
          if k not in kategorie:
            kategorie[k] = []
          kategorie[k].append(key)
    json.dump(kategorie, outfile, default=lambda o: o.__dict__, ensure_ascii=False)

  print("Done")


if __name__ == "__main__":
  main()