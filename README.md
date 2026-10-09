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
- Enne elementide loomist kuvatakse impordi eelvaade: iga sisendis oleva koodi punktide arv ja kooditabelist puuduva koodi juures märge `kood puudub`. `Import` käivitab impordi; `Katkesta` katkestab joonist muutmata. Pikk koodiloend kuvatakse mitmel lehel: `Edasi` avab järgmise lehe ja viimase lehe `Import` käivitab impordi. Nuppude tegevust selgitavat teksti ei kuvata. Sama MsgBox-akna nupunimed muudetakse 32-bitise Windowsi API kaudu; nimede käitustest PowerDraftis on veel tegemata.
- Iga sisendpunkti keskmesse luuakse täitmata ring raadiusega 0,08 m levelile `MOOTMPUNKT`. Punktinumber lisatakse 0,2 m kõrguse ja laiusega tekstina levelile `MOOTNR`, ringi keskmest 0,25 m paremale ja 0,15 m üles. Punktinumbri joondus on `Left Center`.
- Iga punktinumbri alla lisatakse sisendi Z-kõrgus levelile `MOOTKORG-EH2000`, näiteks `1.65`. Kõrgustekst on numbrist 0,3 m allpool, sama X-asukohaga ning sama fondi, suuruse ja joondumisega nagu `MOOTNR`. Väärtusel on kaks komakohta ja punkt kümnenderaldajana; kõrgussüsteemi teisendust ei tehta.
- `RING02` celli puhul lisatakse kõrgus samale levelile ühe tekstina cellist paremale, näiteks `1.65`. Väärtusel on kaks komakohta ja punkt kümnenderaldajana. Teksti kõrgus on 0,85 m, laius 0,65 m ning font `ENGINEERING`.
- Koodide vastavused loetakse failist `kooditabel.csv`. Väljad on **Kood | Koodnimi | Tüüp | Leveli nimi | Joone/Celli nimi | Skaala**; tüüp on `CELL` või `JOON`. `Koodnimi` on ainult informatiivne, võib olla tühi ja importija ignoreerib seda. Tabelis ei tohi olla korduvaid koode.
- CELL-tüübi positiivne mõõtkordaja loetakse väljast `Skaala` ja rakendatakse X- ning Y-suunas; Z-suuna mõõtkordaja jääb 1. Pöördenurk on 0. JOON-tüübi puhul skaalat ei kasutata ja väli võib olla tühi. Ka celli alamelemendid viiakse tabelis määratud levelile.
- JOON-tüübi kirjed loovad eraldi avatud linestring'id tabelis määratud levelile. „Joone/Celli nimi“ peab sisaldama MicroStationis kättesaadava joonestiili nime (näiteks `TEE`) või standardstiili numbrit 0–7. Stiiliobjekt määratakse otse elemendile, mitte ei võeta leveli ByLevel-stiilist.
- Koodi muutus lõpetab joone. Tühje ridu eiratakse; need joont ei katkesta. Sama koodi hilisem uus grupp alustab uut joont.
- Ühe punktiga JOON-kirjest joont ei looda, kuid selle mõõtepunkt ja punktinumber imporditakse. Vigane andmerida katkestab impordi enne geomeetria lisamist. Kooditabelist puuduva koodiga punktile luuakse ainult mõõtepunkt ja punktinumber; puuduvad koodid näidatakse impordi eelvaates. Õnnestunud impordi lõpus infoakent ei kuvata. Impordivea korral kuvatakse veateade.
- Puuduv level luuakse. Programm seob celliteegi aktiivse sessiooniga.
- Korduv käivitus lisab elemendid uuesti. Kasuta esmaseks kontrolliks tühja testjoonist.
- Mõõtepunkti kõrgustekst lisatakse igale sisendpunktile; `RING02` cellile lisatakse lisaks selle eraldi kõrgustekst. Kooditabelit muudetakse CSV-failis; eraldi kasutajaliidest ei ole.

VBA toe viide: https://bentleysystems.service-now.com/community?id=kb_article&sysparm_article=KB0109946

Failidialoog kasutab Windowsi Unicode API-t ja on mõeldud **32-bitisele PowerDraft V8i-le**, mitte 64-bitisele CONNECT-ile. Kasutaja kinnitas v2 toimimist PowerDraftis 6. oktoobril 2026. Kõiki CSV vormingu variante ei ole PowerDraftis eraldi kontrollitud.

## Järgmine etapp

Kooditabeli muutmise kasutajaliides ja celliteegi valitud asukoha salvestamine lisanduvad hiljem.

## Kaks tööriistanuppu

Impordi samasse `ValiImport.mvba` projekti lisaks `ValiImport.bas` moodulile failid `JoonteYhendamine.bas` ja `clsJoinLines.cls` (File → Import File). `clsJoinLines` peab olema klassimoodul. Seejärel tee Debug → Compile ja salvesta projekt.

Määra PowerDrafti tööriistanuppude Key-in väljadele järgmised käsud:

| Nupp | Key-in |
| --- | --- |
| Andmete sisselugemine | `VBA RUN [ValiImport]ValiImport.Import` |
| Joonte ühendamine | `VBA RUN [ValiImport]JoonteYhendamine.Start` |
| Kõrgusarvu pööramine | `VBA RUN [ValiImport]KorgusePooramine.Start` |

Joonte ühendamisel vali esimene joon ühe vasakklõpsuga, seejärel teine joon ühe vasakklõpsuga. Esimene valik tõstetakse esile. Käsu kasutamiseks ei ole vaja eelnevat valikukomplekti.

- Toetatud on aktiivse mudeli tavalised jooned ja linestringid. Kaared, complex string'id, cellide osad ja viitefailide jooned ei kuulu selle käsu sisendisse. Lukustatud joont või levelit ei muudeta.
- Kõik tipud säilivad, vajadusel pööratakse tipujärjestus ümber. Lähim otspunktide paar leitakse XYZ-kauguse järgi. Vahe ühendatakse sirglõiguga; täpselt ühine otspunkt lisatakse üks kord.
- Uus linestring luuakse esimese valitud joone põhjal, kasutades seda atribuudišabloonina: level, värv (ka ByLevel), joonestiil ja joonepaksus pärinevad esimeselt joonelt. Importija värviseade sellele käsule ei rakendu.
- Algne joonepaar asendatakse uue linestringiga. Kui asendamine ebaõnnestub, proovib käsk taastada eemaldatud algjooned ja kuvab vea.
- Pärast ühendamist saab valida järgmise paari. Paremklõps tühistab poolelioleva esimese valiku; kui valikut ei ole, lõpetab käsu.

Hiirekäsu, atribuutide ülekande ja tagasivõtmise käitustest PowerDraft V8i-s on veel tegemata. VBA hiirekäsud kasutavad Bentley dokumenteeritud [IPrimitiveCommandEvents liidest](https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0110097).

## Kõrgusarvu pööramine

Lisa samasse projekti tavamoodul `KorgusePooramine.bas` ja klassimoodul `clsRotateHeight.cls`, seejärel uuenda `ValiImport.bas`. Kui klassifaili import paigutab selle Modules alla, loo Insert → Class Module, määra nimeks `clsRotateHeight` ja kleebi sinna klassifaili sisu alates `Option Explicit` reast. Tee Debug → Compile ning salvesta projekt.

Käivita `VBA RUN [ValiImport]KorgusePooramine.Start`. Vali suure kõrgusarvu tekst levelil `KORGUS-EH2000` ühe vasakklõpsuga. RING02-te eraldi valida ei ole vaja. Tekst märgitakse valituks; hiire liigutamisel näidatakse pööratud koopiat. Teine vasakklõps kinnitab pööramise, paremklõps tühistab eelvaate algteksti muutmata. Pärast kinnitamist või tühistamist saab valida järgmise kõrgusarvu. Paremklõps ilma valikuta lõpetab käsu.

Pööratakse XY-tasandis RING02 celli origini ehk keskpunkti ümber. Teksti asukoht ja suund pöörduvad koos, Z-kõrgus ning kaugus keskpunktist säilivad. Hiire algsuund võetakse esimesest klõpsust; järgnev hiire suuna muutus määrab pöördenurga. Celli ja väikest mõõtepunkti kõrgusteksti ei muudeta.

Uutel importidel salvestatakse suure teksti sisse RING02 elemendi ID nimega XData-seosena. Seos säilib pööramisel ja DGN-i salvestamisel. Varasemalt imporditud teksti puhul taastatakse eeldatav keskpunkt senise nihke `X + 0.3`, `Y - 0.5` ja teksti pöördenurga järgi. Vastava RING02 keskpunkt peab asuma 0,01 m tolerantsi piires ning celli kõrgus peab vastama tekstile kahe komakoha täpsusega. Kui vasteid pole või neid on mitu, kuvatakse viga. Puuduva või vigase salvestatud seose korral teist celli automaatselt ei valita. Varasema teksti seos salvestatakse alles pööramise kinnitamisel.

Nurkade ja vana teksti keskpunkti leidmise testid: `tests/test_rotation_geometry.ps1`. Eelvaate, XData-seose ning kinnitamise/tühistamise käitustest PowerDraftis on veel tegemata.

## Värv

Kõigi uuel impordil loodud elementide, ka cellide alamelementide värv on ByLevel. Varem imporditud elemente see muudatus ei muuda. Värvimuudatuse käitustest PowerDraftis on veel tegemata.
