# Automatiserar kontroll av betyg samt skapande av betygsstatistik

Med detta pythonscript kan du automatisera bort en hel del av betysadministrationen.

## Bakgrund och vad scriptet gör
[Läs artikeln på min sajt för mer information om vad och varför](https://mickekring.se/sa-automatiserar-du-kontroll-av-betyg-samt-skapande-av-betygsstatistik/)
<br />Här finns även en film som visar hela flödet.

## Kom igång

### Du behöver Python version 3.x samt Java installerat för att köra scriptet.

1. Spara filen __main.py__ i en mapp på din dator.
2. Skapa 3 mappar i den mappen, som ska heta __betygskatalog__, __betygskatalog_felsökning__ och __betygskatalog_statistik__ så att det ser ut såhär 

![dir](https://user-images.githubusercontent.com/10948066/202915732-21f504c2-fa41-4c23-947a-76e7a7d86c3b.jpg)

3. Installera [Java](https://www.java.com/sv/download/), som krävs av Tabula-modulen. Starta därefter om datorn.
4. Se till att du har python installerat på din dator och installera även modulerna __xlrd__, __xlwt__, __xlsxwriter__, __tinydb__, __termcolor__, __tabula-py__ och __pandas__. Det gör du genom 'pip3 install' <br />Om du inte har koll på detta, så kommer jag släppa en liten tutorial hur du kommer igång med Python på din dator under kommande veckan.
5. Byt namn på din betygskatalog till __betyg.pdf__ och lägg den i mappen __betygskatalog__
6. Kör scriptet __main.py__ och välj 1 eller 2, det vill säga felsökning eller statistik.

## Frågor
Hör av dig till mig på sociala medier, oftast @mickekring, eller via mail på jag@mickekring.se
