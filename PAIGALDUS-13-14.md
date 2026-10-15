# Elementide asetamine ja eksport

See täiendus lisab kaks käsku olemasolevasse **Default.mvba** projekti. Nuppe ega loendeid pole vaja vormile käsitsi joonistada: kood loob need akna avamisel.

## Paigaldamine PowerDraft V8i-s

1. Salvesta olemasolevast `Default.mvba` failist varukoopia ning ava projekt VBA redaktoris.
2. Uuenda mooduleid **ValiImport** ja **JoonteYhendamine** nende `.bas` failidest. Eemalda vastav vana moodul enne uue importimist, et ei tekiks topeltnimesid.
3. Lisa **File → Import File** kaudu uued standardmoodulid:
   - `ElementAndmed.bas`
   - `TooriistaAknad.bas`
   - `ElementideAsetamine.bas`
   - `ElementideEksport.bas`
4. Olemasolevad `KorgusePooramine`, `clsRotateHeight` ja `clsJoinLines` jäävad vajalikuks.
5. Lisa **Insert → Class Module** kaudu kaks klassi. Määra F4 aknas `(Name)` täpselt **clsPlaceFeature** ja **clsElementKirje**. Kopeeri vastavast `.cls` failist sisu **alates `Option Explicit` reast**. Faili alguses olevaid `VERSION`, `BEGIN`, `END` ja `Attribute` ridu ära kopeeri. Mõlemad peavad asuma **Class Modules** all.
6. Lisa **Insert → UserForm** kaudu kaks tühja vormi. Määra F4 aknas `(Name)` täpselt **frmAsetamine** ja **frmEksport**. Ava kummagi vormi kood **F7** abil ning asenda kogu vaikekood vastava faili sisuga:
   - `frmAsetamine.vba.txt`
   - `frmEksport.vba.txt`
7. Vali **Debug → Compile** ja salvesta `Default.mvba`. Vormide loomisel lisab VBA Microsoft Forms teegi viite. Kui `MSForms` ei ole leitav, kontrolli **Tools → References → Microsoft Forms 2.0 Object Library**.

Failid `.vba.txt` sisaldavad vormi koodi, neid ei impordita standardmoodulitena.

## Käivitamine ja toolbox

Elemendi asetamine:

```text
VBA RUN [Default]ElementideAsetamine.Start
```

Eksport:

```text
VBA RUN [Default]ElementideEksport.Start
```

Kui toolbox vajab projekti eelnevat laadimist, kasuta nupu all ühte rida, näiteks:

```text
VBA LOAD "C:\ProgramData\Bentley\PowerDraft V8i (SELECTseries)\WorkSpace\Projects\vba\Default.mvba";VBA RUN [Default]ElementideEksport.Start
```

Asetamise nupu puhul asenda lõpus `ElementideEksport.Start` nimega `ElementideAsetamine.Start`.

## Asetamine

Vali level ning selle alt kooditabelis olev cell või joon. Loendis kuvatakse kood, Info, tüüp ja celli/joonestiili nimi. Kooditabelit otsitakse aktiivse DGN-i kaustast; vajadusel avaneb failivalik.

- **Cell:** vali celliteek ja klõpsa asetuskohal. Järgmise klõpsuga saab asetada sama celli uuesti. Parem klõps lõpetab käsu.
- **Joon:** vasakud klõpsud määravad tipud. Parem klõps salvestab vähemalt kahe tipuga joone ning lubab alustada järgmist. Veel üks parem klõps lõpetab käsu. Üksiku alguspunkti korral parem klõps tühistab selle.
- Kasutatakse kooditabeli levelit, joonestiili ja celli skaalat. Cell on algselt pööramata. Z tuleb PowerDraftis määratud punktist; AccuSnap on lubatud ja eraldi kõrgusevälja ei ole.
- `RING02` saab ka suure kõrgusteksti ning olemasoleva pööramiskäsu jaoks seose celliga. Käsitsi asetamisel mõõtepunkti ringi ega väikeseid mõõdistustekste juurde ei looda.
- Uus cell ja iga uus joone tipp saavad mudelis salvestatud numbrite hulgast vaba positiivse punktinumbri. Number salvestatakse elemendi andmetesse, eraldi numbriteksti ei joonistata.

## Punktinumbrite seos algandmetega

Uuendatud import salvestab celli juurde algse `pnr` ja koodi; joone juurde iga tipu algse `pnr` ning joone koodi. Algusnullid ja korduvad punktinumbrid säilivad. Punktinumbrit ei asendata DGN-i sisemise elemendi ID-ga.

Kui punktile celli ega joont ei looda (näiteks puuduv kood või üksik JOON-punkt), salvestatakse seos selle mõõtepunkti ringile. Ülejäänud mõõtepunkti ringe ja abitekste ei ekspordita celli või joone kõrval uuesti.

Joonte ühendamine säilitab punktinumbrite järjekorra koos tippude pööramisega. Joone kood tuleb esimeselt valitud joonelt. Täpselt ühine otspunkt väljastatakse üks kord ning säilitab esimese valitud joone otspunkti numbri.

Varasemad impordid pole algseid punktinumbreid elementide külge salvestanud. Nende puhul ei taastata numbreid kõrval asuvatest tekstidest; eksporditavatele elementidele määratakse vabad numbrid ja salvestatakse need.

## Eksport

Käsk loeb aktiivset mudelit, sealhulgas peidetud levelitel olevaid elemente; viitefaile ei loeta. Arvestatakse tavalisi ja shared-celle, jooni, linestring'e, kujundeid ning sirgetest lõikudest koosnevaid complex-jooni ja -kujundeid. Celli sisu ei loeta eraldi objektidena.

1. Vali leitud koodide loendist üks või mitu koodi, näiteks **200** ja **22**. Loendis kuvatakse elementide ja eksporditavate punktide arv. Celli kohta on üks punkt, joone puhul rida iga tipu kohta.
2. Seoseta elemendil otsitakse kood leveli ja celli nime või joonestiili järgi. ByLevel-joone puhul kasutatakse leveli joonestiili. Mitme vaste või puuduva vastega grupid kuvatakse eraldi. Vali selline grupp, vali rippmenüüst kood ja vajuta **Määra kood**. Seejärel vali eksporditavad koodid uuesti.
3. Märgi soovitud veerud **pnr**, **x**, **y**, **z**, **kood**. Vaikimisi on kõik valitud.
4. Vali **Salvesta faili** või **Too ekraanile**. **Ehr** on nähtav nupp, millel praegu tegevust ei ole.

Fail on päisega UTF-8 TXT, komaeraldusega ja punktiga kümnenderaldajana. Ekraanil on kopeeritav tabel tabulaatoritega. X on põhjasuund, Y idasuund; koordinaadid võetakse praegusest geomeetriast ja väljastatakse kolme komakohaga. Märkusega kuues veerg ei kuulu eksporti. Kõigi viie veeruga väljund sobib impordi sisendiks; joonegruppide eraldamiseks pole väljundis eraldi joone ID veergu.

Numbrid ja uued koodiseosed kirjutatakse mudelisse väljundi loomisel, ka ekraanile toomisel. Ainult akna avamine või **Määra kood** ei kirjuta neid veel mudelisse. Failidialoogi katkestamine ei loo väljundit ega salvesta uusi seoseid. Lukustatud elemendi või leveli korral peab uue seose salvestamiseks lukustuse vabastama. DGN tuleb pärast muudatusi tavapäraselt salvestada.

Kaared ja kõveratega complex-elemendid jäetakse välja; nende arv näidatakse aknas. Kui seotud joone tippude arv on vahepeal muutunud, näidatakse seda eraldi veaga grupina: teiste koodide eksport jääb võimalikuks, selle elemendi väljund peatatakse vale `pnr` omistamise vältimiseks. Algse seose taastamiseks tuleb taastada vastav tipujärjestus või element uuesti importida. Seos jälgib tippude järjestust: väljaspool meie ühendamiskäsku tehtud tippude ümberjärjestamine vajab samuti punktiseoste uuendamist.

## Kontrollimine

Automaatkontrollid ei käivita PowerDrafti ega muuda kasutaja DGN-e:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/test_point_data.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/test_join_geometry.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/test_rotation_geometry.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/test_v8i_api.ps1
```

Andmeseoste ja ekspordi kontroll kasutab lähtekoodist loetud päris rutiine ning asendab ainult PowerDrafti objektid testobjektidega. API kontroll loeb kohalikku V8i tüübiteeki. Need ei asenda VBA projekti kompileerimist ega akende ja hiirekäskude proovimist PowerDraftis.

Proovi test-DGN-is importi, kahe koodi valimist, mõlemat väljundit, korduvat eksporti ning kahe imporditud joone ühendamist. Korduval ekspordil peavad `pnr` väärtused säilima. Celli liigutamisel peavad koordinaadid muutuma ja `pnr` samaks jääma.
