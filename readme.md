Käytetyn datan ollessa tutkimuksen alaista, en voi jakaa tässä testaamista varten Excel tiedostoa.
Tehtävänäni oli mitata kaksoistähtien (kataklymisten muuttujien --> kääpiönovien --> SU UMa-kohteiden --> WZ Sge alaluokituksen kohteiden) kirkkauksia prosessoiduista kuvista.
Kasasin mitatut kirkkaudet Excel-taulukoihin, suoritin tarpeelliset laskut ja kasasin lopputulokset alla olevilla scripteillä.

---

# 1. "new_data_processing.bas" sisältää useamman käytettävän funktion:

* "CombineSheets" --> Pääohjelma, jossa luodaan "Processed" välilehti, jolle siirretään kaikista välilehdistä välilehden nimi (kohteen nimi), sekä kyseisen kohteen kirkkaudet eri filttereissä virheineen.

* "generate_scatterplot" -> Luo "Processed" välilehteen siirretystä aineistosta pistekaavion huomioiden virheet

* "generate_files" --> Luo "Processed" välilehteen siirretystä aineistosta tekstitiedostot kohteen nimen mukaisesti

* "change_error_columns" --> Sen sijaan, että "Processed" välilehdessä kohteiden kirkkauksien virheenä olisi matemaattisesti määritelty virhe, otetaan harkinnan kautta todettu virhe "kommentti"-soluista.

---

# 2. "other_sources_processing.bas" sisältää "Former_magnitudes" scriptin, jolla eri lähteistä (SkyMapper, Pan-Starrs, SDSS, DES...) etsityt (ja muunnetut) kirkkaudet siirrettiin "Historical" välilehteen.
