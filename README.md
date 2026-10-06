# Väliandmete import – V8i prototüüp v2

Töötav VBA prototüüp MicroStation PowerDraft V8i 08.11.09.829 jaoks. Kasutaja kinnitas failivaliku, SPYMAR celli ja kahe joone loomise õigetele levelitele. Repo sisaldab lähtemoodulit (.bas); .mvba projekt salvestatakse PowerDraftis.

1. Ava PowerDraft V8i-s tühi **3D DGN**, mille põhiühik on **meeter**.
2. Ava Utilities → Macro → Project Manager (või sisesta `VBA PROJECT MANAGER`). Loo uus projekt nimega **ValiImport** ja salvesta see .mvba failina.
3. Ava projekti VBA redaktor. Kui vana `ValiImport` moodul on juba olemas, eemalda see Project Exploreris paremklõpsuga **Remove ValiImport** (soovi korral ekspordi varukoopia). Vali **File → Import File** ja impordi uus `ValiImport.bas`. Ära jäta vana ja uut moodulit korraga projekti.
4. Vali **Debug → Compile** ning salvesta projekt.
5. Käivita makro `ValiImport.Import` (või sisesta `VBA RUN [ValiImport]ValiImport.Import`).
6. Vali avanevas failidialoogis oma **TXT või CSV**. Failinimi võib olla suvaline. Kui `parnu_tm_mkm.cel` on samas kaustas, kasutatakse seda; muidu avaneb celliteegi valimise dialoog. Cancel katkestab ilma importimata.
7. Tee Fit View. Kontrolli celli suurust, asukohta ja leveleid.

Oodatav tulemus: üks SPYMAR cell levelil ALUSVORK ning kaks eraldi kahepunktijoont levelil HOONE. Esimese celli asukoht joonisel: X=529317.082, Y=6470412.167, Z=3.801.

## Prototüübi reeglid

- Sisend: `pnr,x,y,z,kood`. Päis on lubatud. Eraldajaks sobib koma, semikoolon või tabulaator. Semikooloni/tabulaatori puhul võib kümnendmärk olla koma. Toetatud on jutumärkides väljad ja Exceli `sep=;` algusrida.
- Toetatud on UTF-8 (ka BOM-iga), BOM-iga UTF-16 ja lihtsad ASCII andmeread. Päise nimed peavad olema pnr, x, y, z, kood.
- Algse näidisfaili `Selgitus:` rida lõpetab andmete lugemise; sellele järgnev selgitus ei ole imporditav andmestik. Muu vigane rida annab vea koos rea sisuga.
- Faili X on põhi, Y ida. Joonise X saab faili Y; joonise Y saab faili X. Z säilib.
- Kood 2: SPYMAR, ALUSVORK. Mõõtkordaja 1 ja pöördenurk 0; sobiv suurus tuleb joonisel kontrollida. Ka celli alamelemendid viiakse ALUSVORK levelile.
- Koodid 6 ja 10: eraldi avatud linestring'id, HOONE. Esialgu tavaline joonekuju, aktiivse seadistuse joonesümboolikaga.
- Koodi muutus lõpetab joone. Tühje ridu eiratakse; need joont ei katkesta. Sama koodi hilisem uus grupp alustab uut joont.
- Alla kahe punktiga joonegrupp, tundmatu kood või vigane andmerida katkestab impordi enne geomeetria lisamist.
- Puuduv level luuakse. Programm seob celliteegi aktiivse sessiooniga.
- Korduv käivitus lisab elemendid uuesti. Kasuta esmaseks kontrolliks tühja testjoonist.
- Punktinumbreid ja kõrgustekste ei joonistata. Puudub kooditabeli kasutajaliides.

VBA toe viide: https://bentleysystems.service-now.com/community?id=kb_article&sysparm_article=KB0109946

Failidialoog kasutab Windowsi Unicode API-t ja on mõeldud **32-bitisele PowerDraft V8i-le**, mitte 64-bitisele CONNECT-ile. Kasutaja kinnitas v2 toimimist PowerDraftis 6. oktoobril 2026. Kõiki CSV vormingu variante ei ole PowerDraftis eraldi kontrollitud.

## Järgmine etapp

Kooditabeli väljad: **Kood | Koodnimi | Tüüp | Leveli nimi | Joone/Celli nimi**.

Praegune moodul kasutab veel koodis määratud vastavusi (2, 6, 10). Eraldi kooditabeli lugemine ja muutmise kasutajaliides on järgmine arendusetapp. Celliteegi valitud asukoha salvestamine lisandub hiljem.
