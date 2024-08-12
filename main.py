# Standard library
import datetime
from pathlib import Path
from helperFunctions import sqlfiler, EmailObjekt

MODTAGER: str = 'Rasmus.steffen.bendtsen@regionh.dk'

FILNAVN: str = 'Evalueringsansættelser'

BODYTXT: str ='''Kære Carsten.

Her er seneste oversigt over sygeplejersker ansat i evalueringsansættelser

Venlig hilsen
Data og Digitalisering
'''


# sendes først kommende mandag efter d. 9 i hver måned
# hvis dags dato er mellem d. 10 og d. 16 og det er mandag
if 10 <= datetime.datetime.now().day < 17 and datetime.datetime.now().weekday() == 0:

    sqlfiler(f"{FILNAVN}.sql", firstHasOutput=True).toExcel()

    mail = EmailObjekt(
                modtager = MODTAGER,
                emne = FILNAVN,
                bodytxt = BODYTXT,
                attachments = Path.joinpath(Path.cwd(),f"{FILNAVN}.xlsx"))

    mail.SendMail(DisplayBeforeSend=False)
    print(f'Sender {FILNAVN} leverance til {mail.modtager}')
else:
    print(f'{FILNAVN} leverance skal ikke kører idag.')