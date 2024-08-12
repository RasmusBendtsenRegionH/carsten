from pathlib import Path
import polars as pl
import pyodbc
import win32com.client

class sqlfiler():


    """
    Anvendes til at processere SQL filer
    """

    def __init__(self, *filnavn: str, firstHasOutput: bool = True):
        """
        Indlæser SQL filer

        Examples:
            >>> sqlfiler(fil.sql , firstHasOutput=True)
            >>> sqlfiler(temptabeller.sql , fil.sql, firstHasOutput=False)

        Args:
            filnavn (str): SQL filer inklusiv filendelse.
            firstHasOutput (bool, optional): Anvendes hvis den første query opsætter temp tabeller, i så fald sæt til False
        """
        self.filnavn = filnavn
        self.firstHasOutput = firstHasOutput

    def processqueries(self) -> dict:
        """
        Omdanner queries til dataframes. Hvis firstHasOuput = False, så returnes ingen dataframe for den første query.

        Examples:
            >>> queries = sqlfiler(fil.sql, firstHasOutput=True) 
            >>> queries.processqueries()
            {"fil":pl.DataFrame}


        Returns:
            (dict): key = queryens navn, value = queryens output som pl.dataframe
        """

        def readfile(filnavn):
            """Indlæser SQL script"""
            fd = open(filnavn, 'r')
            sqlFile = fd.read()
            fd.close()
            return sqlFile

        def executeQuery(self) -> pl.DataFrame | None:
            """
            Kører query gennem P01

            Args:
                query (str): SQL script
                toDataframe (bool): Skal scriptets output omdannes til Dataframe

            Returns:
                pl.DataFrame | None: Dataframe hvis toDataframe = True, ellers så køres koden uden at returnere noget, anvendes til oprettelse af Temp tabeller
            """
            conn = pyodbc.connect('Driver={SQL Server};'
                                  'Server=RGHSQLCOKP01;'
                                  'Database=LON_HR;'
                                  'Integrated Security=true'
                                  )
            # Hvis SQL filerne er ikke UFT-8 encoded, så udkommentar linje under.
            conn.setencoding("Windows-1252")
            cursor = conn.cursor()
            if self.firstHasOutput:
                return {filnavn[:-4]: pl.read_database(readfile(filnavn), cursor) for filnavn in self.filnavn}
            else:
                cursor.execute(readfile(self.filnavn[0]))
                return {filnavn[:-4]: pl.read_database(readfile(filnavn), cursor) for filnavn in self.filnavn[1:]}

        return executeQuery(self)


    def toExcel(self,placering: Path = Path.cwd()) -> None:
        """
        Eksportere dataframes til CSV

        Examples:
            >>> queries = sqlfiler("fil.sql", firstHasOutput=True).toExcel()
            >>> DataFrame gemmes som xlsx i placering. CWD som default

        Args:
            placering (Path): Hvor skal filerne gemmes (Anvend Pathlib)

        Returns:
            (None): Filer bliver gemt i [placering]
        """
        
        for k,v in self.processqueries().items():

            v.write_excel(Path.joinpath(
                placering, f"{k}.xlsx"))
        return None

class EmailObjekt:

    """
    Anvendes som skabelon, til at sende emails fra Outlook gennem Python
    """

    def __init__(self, 
                modtager: str,
                emne: str, 
                afsender:str = 'analyse.center-for-hr@regionh.dk',
                bodytxt: str = None, 
                CC='analyse.center-for-hr@regionh.dk',
                category = 'Faste leverance',
                **attachments: str) -> None:

        """De forskellige oplysninger som er nødvendige for at kunne lave en mail
        Args:
            afsender (str): Afsender på mailen. Defaults to 'analyse.center-for-hr@regionh.dk'.
            emne (str): Emnefelt
            bodytxt (str, optional): Teksten i selve mailen. Defaults to None.
            CC (str, optional): CC på mailen. Defaults to 'analyse.center-for-hr@regionh.dk'.
            category (str, optional): CC på mailen. Defaults to 'Faste leverancer'.
            modtager (str): mailadresse 
            attachments (str, optional): Hver vedhæftningen skal være den fulde filsti inklusiv filnavn og endelse
        """
        self.afsender = afsender
        self.emne = emne
        self.bodytxt = bodytxt
        self.CC = CC
        self.category = category
        self.modtager = modtager
        self.attachments = attachments

    def SendMail(self, DisplayBeforeSend=True) -> None:
        """Anvendes til at sende mailen
        Args:
            DisplayBeforeSend (bool, optional): Skal mailen vises først, så man kan validere eller sendes med det samme. Defaults to True.
        """
        outlook = win32com.client.Dispatch('outlook.application')
        newmail = outlook.CreateItem(0)
        newmail.SentOnBehalfOfName = self.CC
        newmail.Subject = self.emne
        # wincom anvender ";" som seperator mellem mail adresserne. 
        newmail.To = self.modtager
        newmail.CC = self.CC
        if self.attachments:  # hvis filen har attachments
                newmail.attachments.Add(str(self.attachments['attachments']))
        newmail.Body = self.bodytxt
        newmail.Categories = self.category
        if DisplayBeforeSend:
            newmail.Display(True)
        else:
            newmail.Send()
