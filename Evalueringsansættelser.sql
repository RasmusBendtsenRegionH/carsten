WITH
    doks
    as
    (
        SELECT CPR, Oprettelsestidspunkt, Dokumenttitel, dokumenter.LeveranceDato
        FROM [LON_HR].[SD_SYSTEM_INFO].[V_SD_DOKUMENTER] dokumenter

            -- ikke alle dokumenter har en tilhørende skabelon
            LEFT JOIN [LON_HR].[SD_SYSTEM_INFO].[V_SD_DOKUMENTTYPER_SKABELONER] skabeloner ON skabeloner.ID = dokumenter.dokumenttype
        WHERE skabeloner.navn = '1.01 Ansættelsesbrev - Sygeplejerske i evalueringsansættelse'
    )


SELECT DISTINCT
    Person.navn AS Navn,
    Person.tjnr AS [TJNR],
    Person.ANSDATO AS [Ansættelsesdato],
    person.INST,
    org.org3txt AS Afdeling,
    org.org2txt AS Afsnit

FROM
    [LON_HR].[SD].[SD_PERSON] Person
    LEFT JOIN SD.SD_DIM_ORG org ON 1 = 1
        AND org.CURRENT_ROW = 1
        AND org.INST = Person.INST
        AND org.AFD = Person.AFD
        AND org5txt = 'drift'
    LEFT JOIN SD.SD_DIM_STILLINGSKODE stil ON 1 = 1
        AND stil.CURRENT_ROW = 1
        AND stil.INST = '8H'
        AND stil.STILKO0 = Person.STILKO
    INNER JOIN doks on doks.CPR = person.CPR

WHERE 1=1

    AND person.STAT in('0','1','3')
    AND CONVERT(date,getdate()) BETWEEN Person.[START] AND Person.SLUT

order BY TJNR
