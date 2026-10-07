# Väliandmete import – V8i prototüüp v2

Töötav VBA prototüüp MicroStation PowerDraft V8i 08.11.09.829 jaoks. Kasutaja kinnitas failivaliku, SPYMAR celli ja kahe joone loomise õigetele levelitele. Repo sisaldab lähtemoodulit (.bas); .mvba projekt salvestatakse PowerDraftis.

1. Ava PowerDraft V8i-s tühi **3D DGN**, mille põhiühik on **meeter**.
2. Ava Utilities → Macro → Project Manager (või sisesta `VBA PROJECT MANAGER`). Loo uus projekt nimega **ValiImport** ja salvesta see .mvba failina.
3. Ava projekti VBA redaktor. Kui vana `ValiImport` moodul on juba olemas, eemalda see Project Exploreris paremklõpsuga **Remove ValiImport** (soovi korral ekspordi varukoopia). Vali **File → Import File** ja impordi uus `ValiImport.bas`. Ära jäta vana ja uut moodulit korraga projekti.
4. Vali **Debug → Compile** ning salvesta projekt.
5. Käivita makro `ValiImport.Import` (või sisesta `VBA RUN [ValiImport]ValiImport.Import`).
6. Vali avanevas failidialoogis oma **TXT või CSV**. Failinimi võib olla suvaline. Kui `kooditabel.csv` ja `parnu_tm_mkm.cel` on samas kaustas, kasutatakse neid; muidu avaneb vastava faili valimise dialoog. Cancel katkestab ilma importimata.
7. Tee Fit View. Kontrolli celli suurust, asukohta ja leveleid.

Oodatav tulemus: üks SPYMAR cell levelil ALUSVORK, kaks eraldi kahepunktijoont levelil HOONE ning iga sisendpunkti juures mõõtepunkti ring ja punktinumber. Esimese celli asukoht joonisel: X=529317.082, Y=6470412.167, Z=3.801.

## Prototüübi reeglid

- Sisend: `pnr,x,y,z,kood`. Päis on lubatud. Esimene sisuline rida võib selle asemel sisaldada töö nime, kuupäeva või muud faili infot: kui see ei ole päis ega viieväljaline andmerida, jäetakse see vahele. Korrektne viieväljaline esimene punktirida imporditakse tavaliselt. Eraldajaks sobib koma, semikoolon või tabulaator. Semikooloni/tabulaatori puhul võib kümnendmärk olla koma. Toetatud on jutumärkides väljad ja Exceli `sep=;` algusrida.
- Toetatud on UTF-8 (ka BOM-iga), BOM-iga UTF-16 ja lihtsad ASCII andmeread. Päise nimed peavad olema pnr, x, y, z, kood.
- Algse näidisfaili `Selgitus:` rida lõpetab andmete lugemise; sellele järgnev selgitus ei ole imporditav andmestik. Muu vigane rida annab vea koos rea sisuga.
- Faili X on põhi, Y ida. Joonise X saab faili Y; joonise Y saab faili X. Z säilib.
- Iga sisendpunkti keskmesse luuakse täitmata ring läbimõõduga 0,2 m levelile `MOOTMPUNKT`. Punktinumber lisatakse 0,1 m kõrguse ja laiusega tekstina levelile `MOOTNR`, ringi keskmest 0,15 m paremale ja üles.
- `RING02` celli puhul lisatakse kõrgus samale levelile ühe tekstina cellist paremale, näiteks `1.65`. Väärtusel on kaks komakohta ja punkt kümnenderaldajana. Teksti kõrgus on 0,85 m, laius 0,65 m ning font `ENGINEERING`.
- Koodide vastavused loetakse failist `kooditabel.csv`. Väljad on **Kood | Koodnimi | Tüüp | Leveli nimi | Joone/Celli nimi | Skaala**; tüüp on `CELL` või `JOON`. `Koodnimi` on ainult informatiivne, võib olla tühi ja importija ignoreerib seda. Tabelis ei tohi olla korduvaid koode.
- CELL-tüübi positiivne mõõtkordaja loetakse väljast `Skaala` ja rakendatakse X- ning Y-suunas; Z-suuna mõõtkordaja jääb 1. Pöördenurk on 0. JOON-tüübi puhul skaalat ei kasutata ja väli võib olla tühi. Ka celli alamelemendid viiakse tabelis määratud levelile.
- JOON-tüübi kirjed loovad eraldi avatud linestring'id tabelis määratud levelile. „Joone/Celli nimi“ peab sisaldama MicroStationis kättesaadava joonestiili nime (näiteks `TEE`) või standardstiili numbrit 0–7. Stiiliobjekt määratakse otse elemendile, mitte ei võeta leveli ByLevel-stiilist.
- Koodi muutus lõpetab joone. Tühje ridu eiratakse; need joont ei katkesta. Sama koodi hilisem uus grupp alustab uut joont.
- Ühe punktiga JOON-kirjest joont ei looda, kuid selle mõõtepunkt ja punktinumber imporditakse. Vigane andmerida katkestab impordi enne geomeetria lisamist. Kooditabelist puuduva koodiga punktile luuakse ainult mõõtepunkt ja punktinumber; pärast importi kuvatakse puuduvate koodide koondhoiatus.
- Puuduv level luuakse. Programm seob celliteegi aktiivse sessiooniga.
- Korduv käivitus lisab elemendid uuesti. Kasuta esmaseks kontrolliks tühja testjoonist.
- Kõrgustekst lisatakse ainult `RING02` cellile. Kooditabelit muudetakse CSV-failis; eraldi kasutajaliidest ei ole.

VBA toe viide: https://bentleysystems.service-now.com/community?id=kb_article&sysparm_article=KB0109946

Failidialoog kasutab Windowsi Unicode API-t ja on mõeldud **32-bitisele PowerDraft V8i-le**, mitte 64-bitisele CONNECT-ile. Kasutaja kinnitas v2 toimimist PowerDraftis 6. oktoobril 2026. Kõiki CSV vormingu variante ei ole PowerDraftis eraldi kontrollitud.

## Järgmine etapp

Kooditabeli muutmise kasutajaliides ja celliteegi valitud asukoha salvestamine lisanduvad hiljem.
