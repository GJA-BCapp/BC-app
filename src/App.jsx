import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, LineChart, Line, Legend } from 'recharts';
import {
  Users, PiggyBank, Wallet, TrendingUp, TrendingDown, Plus, X, Search,
  Download, Calendar, Pencil, Trash2, Baby, FileSpreadsheet, LayoutDashboard,
  Filter, ArrowUpRight, ArrowDownRight, Grid3x3, ListChecks, Landmark, Printer,
  ClipboardList, ArrowLeft, CheckSquare, Square, FileText, Link2,
  Settings, History, RotateCcw, DatabaseBackup, CheckCircle2, AlertCircle, Palette, Clock3, Upload, Phone, MessageCircle, ChevronUp, ChevronDown, Copy, Lock, Sliders
} from 'lucide-react';
import * as XLSX from 'xlsx';

/* ============================== SEED DATA ==============================
   Overgenomen uit de bestaande Excel-administratie van BladelsCreatief
   (ledenlijst_2025_ev.xlsx & Financiele_administratie_BladelsCreatief.xlsx).
   Bij eerste gebruik wordt dit ingeladen; daarna leeft alles in opgeslagen
   staat en kun je alles hier los van bewerken.
========================================================================= */
const SEED_MEMBERS = [{"id":1,"voornaam":"Andrea","tussenvoegsel":null,"achternaam":"Boullart","naam":"Andrea Boullart","email":"aceboullart@gmail.com","adres":"Bogerd 41","postcode":"5521 RR","woonplaats":"Eersel","telefoon":"06-40206454","gebdatum":"1971-12-29","lidsinds":"2025-02-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":2,"voornaam":"Angelina","tussenvoegsel":"van","achternaam":"Kemenade","naam":"Angelina van Kemenade","email":"calenzia@gmail.com","adres":"Aangelag 16","postcode":"5541 GK","woonplaats":"Reusel","telefoon":"06-24479731","gebdatum":"1982-04-23","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":"Bestuurslid","dagdelen":["di 10.00 - 16.00","wo 09.00 - 12.30"]},{"id":3,"voornaam":"Anneke","tussenvoegsel":null,"achternaam":"Lablans","naam":"Anneke Lablans","email":"annekelablans@gmail.com","adres":"Doolandweg 55","postcode":"5531 PL","woonplaats":"Bladel","telefoon":"06-21628015","gebdatum":"1949-04-25","lidsinds":null,"eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":4,"voornaam":"Anja","tussenvoegsel":"van","achternaam":"Hattem","naam":"Anja van Hattem","email":"anjavanhattem@ziggo.nl","adres":"Sint Janstraat 1","postcode":"5525 BH","woonplaats":"Duizel","telefoon":"06-21186177","gebdatum":"1960-09-14","lidsinds":"2026-05-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["di 10.00 - 16.00"]},{"id":5,"voornaam":"Annie","tussenvoegsel":"van","achternaam":"Oorschot","naam":"Annie van Oorschot","email":"annievanoorsc hot1@gmail.com","adres":"Bertus Aafjeshof 137","postcode":"5531 SV","woonplaats":"Bladel","telefoon":"06-12559325","gebdatum":"1958-09-12","lidsinds":"2025-12-04","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":6,"voornaam":"Anton","tussenvoegsel":null,"achternaam":"Cools","naam":"Anton Cools","email":"anton.cools@gmail.com","adres":"Claassenpark 4","postcode":"5527 BT","woonplaats":"Hapert","telefoon":"06-82607643","gebdatum":"1952-09-03","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":7,"voornaam":"Arianne","tussenvoegsel":"van","achternaam":"Hout","naam":"Arianne van Hout","email":"ariannevanhout@gmail.com","adres":"De Loop 6","postcode":"5501 ER","woonplaats":"Veldhoven","telefoon":"06-18482062","gebdatum":"1951-05-24","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":8,"voornaam":"Brigitte","tussenvoegsel":null,"achternaam":"Burgmans-Lievens","naam":"Brigitte Burgmans-Lievens","email":"brigitte_burgmans@hotmail.com","adres":"Hagelkruis 8","postcode":"5571 PC","woonplaats":"Bergeijk","telefoon":"06-12267952","gebdatum":"1969-06-23","lidsinds":"2025-01-01","eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":9,"voornaam":"Bonnie","tussenvoegsel":null,"achternaam":"Boelhouwers","naam":"Bonnie Boelhouwers","email":"bonnieb1@hotmail.com","adres":"Godfried Bomanslaan 11","postcode":"5531 VH","woonplaats":"Bladel","telefoon":"06-41061405","gebdatum":"2000-08-23","lidsinds":"2026-03-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 19.00 - 22.00"]},{"id":10,"voornaam":"Cor","tussenvoegsel":"van","achternaam":"Gool","naam":"Cor van Gool","email":"corvgool@outlook.com","adres":"Roodborstje 9","postcode":"5531 LM","woonplaats":"Bladel","telefoon":"06-22668215","gebdatum":"1954-01-27","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":"Voorzitter","dagdelen":["di 10.00 - 16.00","wo 09.00 - 12.30"]},{"id":11,"voornaam":"Cora","tussenvoegsel":"van","achternaam":"Hoorn","naam":"Cora van Hoorn","email":"coradorsplace@gmail.com","adres":"Willem Klooslaan 22","postcode":"5531 TN","woonplaats":"Bladel","telefoon":"06-26781715","gebdatum":"1948-08-23","lidsinds":"2026-05-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 19.00 - 22.00"]},{"id":12,"voornaam":"Daan","tussenvoegsel":"van den","achternaam":"Enden","naam":"Daan van den Enden","email":"daanenmarja@gmail.com","adres":"Kersenerf 8","postcode":"5505 LD","woonplaats":"Veldhoven","telefoon":"06-15479992","gebdatum":"1958-03-05","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 19.00 - 22.00"]},{"id":13,"voornaam":"Dick","tussenvoegsel":null,"achternaam":"Scheepmaker","naam":"Dick Scheepmaker","email":"dick.scheepmaker@icloud.com","adres":"Beemke 67","postcode":"5534 AG","woonplaats":"Netersel","telefoon":"06-18935699","gebdatum":"1933-04-20","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":14,"voornaam":"Dirk","tussenvoegsel":"van","achternaam":"Loon","naam":"Dirk van Loon","email":"info@dirkvloon.nl","adres":"Turnhoutseweg 20","postcode":"5541 NK","woonplaats":"Reusel","telefoon":"06-20796795","gebdatum":null,"lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":15,"voornaam":"Dorien","tussenvoegsel":null,"achternaam":"Reijnders","naam":"Dorien Reijnders","email":"dorienvoetverzorging@hotmail.com","adres":"Het Aangelag 3","postcode":"5531 XK","woonplaats":"Bladel","telefoon":"06-53506018","gebdatum":"1957-06-27","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":16,"voornaam":"Elle","tussenvoegsel":null,"achternaam":"Spooren","naam":"Elle Spooren","email":"iagm.spooren@outlook.com","adres":"Smitseind 34","postcode":"5525 AP","woonplaats":"Duizel","telefoon":"06-51981746","gebdatum":"1947-09-22","lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":17,"voornaam":"Elly","tussenvoegsel":null,"achternaam":"Hovens","naam":"Elly Hovens","email":"ellyhovens52@gmail.com","adres":"Molenweg 9","postcode":"5531 PN","woonplaats":"Bladel","telefoon":"06-27251201","gebdatum":"1952-01-10","lidsinds":"2025-10-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":18,"voornaam":"Elly","tussenvoegsel":null,"achternaam":"Verrijt","naam":"Elly Verrijt","email":"info@ellyverrijt.nl","adres":"Loo 85","postcode":"5571 KP","woonplaats":"Bergeijk","telefoon":"06-45680686","gebdatum":"1969-09-10","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 19.00 - 22.00"]},{"id":19,"voornaam":"Gerlie","tussenvoegsel":null,"achternaam":"Schel-Tijsen","naam":"Gerlie Schel-Tijsen","email":"joop1@onsbrabantnet.nl","adres":"Van Herlaerhof 32","postcode":"5553 EM","woonplaats":"Valkenswaard","telefoon":"06-17186957","gebdatum":"1962-03-03","lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":20,"voornaam":"Fred","tussenvoegsel":null,"achternaam":"Bleijs","naam":"Fred Bleijs","email":"gjm.bleijs@gmail.com","adres":"Zeegstraat 48","postcode":"5541 EX","woonplaats":"Reusel","telefoon":"06-10485584","gebdatum":"1956-12-09","lidsinds":"2026-05-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":21,"voornaam":"Gerthy","tussenvoegsel":null,"achternaam":"Fleskens","naam":"Gerthy Fleskens","email":"gfleskens@outlook.com","adres":"Frederik Hendrikstraat 42","postcode":" 5502 TH","woonplaats":"Veldhoven","telefoon":"06-24546492","gebdatum":"1942-07-01","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":22,"voornaam":"Gijsbert","tussenvoegsel":null,"achternaam":"Jansen","naam":"Gijsbert Jansen","email":"gcmjansen@outlook.com","adres":"Postelseweg 151","postcode":"5521 RD","woonplaats":"Eersel","telefoon":"06-23524363","gebdatum":"1960-06-19","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":"Penningmeester","dagdelen":["wo 19.00 - 22.00"]},{"id":23,"voornaam":"Guus","tussenvoegsel":null,"achternaam":"Goossens","naam":"Guus Goossens","email":"apj.goossens@hccnet.nl","adres":"De Voren 8","postcode":"5527 HX","woonplaats":"Hapert","telefoon":"06-23524743","gebdatum":"1952-08-19","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":24,"voornaam":"Hanneke","tussenvoegsel":"de","achternaam":"Groot","naam":"Hanneke de Groot","email":"hanneke.degroot@gmail.com","adres":"Dijckmeesterstraat 3","postcode":"5528 AN","woonplaats":"Hoogeloon","telefoon":"06-44804308","gebdatum":"1953-06-26","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":"Bestuurslid","dagdelen":["wo 19.00 - 22.00"]},{"id":25,"voornaam":"Harry","tussenvoegsel":null,"achternaam":"Jacobs","naam":"Harry Jacobs","email":"jacobs.ha5@gmail.com","adres":"De kerkschuur 16","postcode":"5087 BZ","woonplaats":"Diessen","telefoon":"06-51026952","gebdatum":"1953-04-06","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":26,"voornaam":"Hennie","tussenvoegsel":null,"achternaam":"Walenberg","naam":"Hennie Walenberg","email":"hennie-walenberg@hotmail.com","adres":"Het Hofveld 18","postcode":"5531 GL","woonplaats":"Bladel","telefoon":"06-19085866","gebdatum":"1954-11-04","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":27,"voornaam":"Henny","tussenvoegsel":null,"achternaam":"Herps","naam":"Henny Herps","email":"herpshenny@gmail.com","adres":"Mollenstraat 26","postcode":"5571 BL","woonplaats":"Bergeijk","telefoon":"06-11285209","gebdatum":"1952-05-14","lidsinds":"2025-01-01","eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":28,"voornaam":"Hettie","tussenvoegsel":null,"achternaam":"Damen","naam":"Hettie Damen","email":"hejada77@hotmail.com","adres":"Arnold van Rodelaan 77","postcode":"5527 BS","woonplaats":"Hapert","telefoon":"06-40439584","gebdatum":"1950-01-24","lidsinds":null,"eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":29,"voornaam":"Huub ","tussenvoegsel":null,"achternaam":"Kolsters","naam":"Huub  Kolsters","email":"h.kolsters2@upcmail.nl","adres":"Sniederslaan 140","postcode":"5531 EN","woonplaats":"Bladel","telefoon":"0497-360714","gebdatum":null,"lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":30,"voornaam":"Huub","tussenvoegsel":"ter ","achternaam":"Hart","naam":"Huub ter  Hart","email":"huubterhart@gmail.com","adres":"Kervelstraat 47","postcode":"5571 HZ","woonplaats":"Bergeijk","telefoon":"06-44126555","gebdatum":null,"lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":31,"voornaam":"Huub","tussenvoegsel":"van de","achternaam":"Krogt","naam":"Huub van de Krogt","email":"huubhvdkrogt@kpnplanet.nl","adres":"Hofstad 56","postcode":"5531 GD","woonplaats":"Bladel","telefoon":"06-22446563","gebdatum":"1948-12-02","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["di 10.00 - 16.00"]},{"id":32,"voornaam":"Ingrid M.","tussenvoegsel":"de","achternaam":"Wit-Baerselman","naam":"Ingrid M. de Wit-Baerselman","email":"madelondewit@hotmail.com","adres":"Postakkers 9","postcode":"5521 AR","woonplaats":"Eersel","telefoon":"06-24570752","gebdatum":"1942-02-02","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":33,"voornaam":"Ingrid","tussenvoegsel":null,"achternaam":"Spliethof","naam":"Ingrid Spliethof","email":"ingridspliethof@gmail.com","adres":"Wijenhof 2","postcode":"5096 CK","woonplaats":"Hulsel","telefoon":"06-46762816","gebdatum":"1974-03-03","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 19.00 - 22.00"]},{"id":34,"voornaam":"Ivo","tussenvoegsel":null,"achternaam":"Strouken","naam":"Ivo Strouken","email":"info@plastifex.nl","adres":"Hofstad 68","postcode":"5531 GD","woonplaats":"Bladel","telefoon":"06-51408314","gebdatum":"1991-08-09","lidsinds":null,"eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":["wo 19.00 - 22.00","do 19.00 - 22.00"]},{"id":35,"voornaam":"Jan","tussenvoegsel":"van ","achternaam":"Buul","naam":"Jan van  Buul","email":"j.buul73@upcmail.nl","adres":"Nieuwstraat 6","postcode":"5527 AT","woonplaats":"Hapert","telefoon":"0497-383206","gebdatum":"1942-07-28","lidsinds":null,"eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":36,"voornaam":"Jeanette","tussenvoegsel":null,"achternaam":"Graamans-Ouberg","naam":"Jeanette Graamans-Ouberg","email":"jeanette_graamans@hotmail.com","adres":"Postels Huufke 22","postcode":"5512 AV","woonplaats":"Vessem","telefoon":"06-22160497","gebdatum":"1949-03-11","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":37,"voornaam":"Jeanne","tussenvoegsel":"van","achternaam":"Leeuwen","naam":"Jeanne van Leeuwen","email":"vanleeuwenkox@gmail.com","adres":"Vendelierstraat 20","postcode":"5525 BR","woonplaats":"Duizel","telefoon":"06-22507839","gebdatum":"1950-07-28","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":38,"voornaam":"Jolanda","tussenvoegsel":null,"achternaam":"Spaan","naam":"Jolanda Spaan","email":"jolandaspaan@hotmail.com","adres":"Kraanvogelweg 17","postcode":"5521 VX","woonplaats":"Eersel","telefoon":"06-22497415","gebdatum":"1948-08-10","lidsinds":"2025-06-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":39,"voornaam":"Joyce","tussenvoegsel":"van der ","achternaam":"Kam","naam":"Joyce van der  Kam","email":"joycevdkam@gmail.com","adres":"Het Aangelag 17","postcode":"5531 XK","woonplaats":"Bladel","telefoon":"06-58933225","gebdatum":"1964-11-28","lidsinds":"2025-08-01","eindelidmaat":"2025-09-18","status":"inactief","functie":null,"dagdelen":[]},{"id":40,"voornaam":"Karien","tussenvoegsel":"van der","achternaam":"Heijden","naam":"Karien van der Heijden","email":"tantepollewob@hotmail.com","adres":"Klokstaart 20","postcode":"5521 WV","woonplaats":"Eersel","telefoon":"06-14391943","gebdatum":"1962-10-11","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["di 10.00 - 16.00"]},{"id":41,"voornaam":"Karin","tussenvoegsel":null,"achternaam":"Tonneijk","naam":"Karin Tonneijk","email":"kbtonneijk@kpnmail.nl","adres":"Reijenburg 16","postcode":"5501 LC","woonplaats":"Veldhoven","telefoon":"06-17308810","gebdatum":"1969-01-31","lidsinds":"2025-10-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":42,"voornaam":"Laura","tussenvoegsel":"van","achternaam":"Dijk","naam":"Laura van Dijk","email":"lauravd99@hotmail.com","adres":"Roeststraat 15","postcode":2370,"woonplaats":"Arendonk, België","telefoon":"0032-468188950","gebdatum":"1999-06-25","lidsinds":"2025-12-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 19.00 - 22.00","do 19.00 - 22.00"]},{"id":43,"voornaam":"Lein","tussenvoegsel":null,"achternaam":"Kloet","naam":"Lein Kloet","email":"leinkloet@gmail.com","adres":"Stijn Streuvellaan 15","postcode":"5531VA","woonplaats":"Bladel","telefoon":"06-12101659","gebdatum":"1938-06-14","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["di 10.00 - 16.00"]},{"id":44,"voornaam":"Liesbeth","tussenvoegsel":null,"achternaam":"Adams","naam":"Liesbeth Adams","email":"liesbeth.adams13@gmail.com","adres":"Boogschutter 16","postcode":"5527 CV","woonplaats":"Hapert","telefoon":"06-17391024","gebdatum":"1965-06-18","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":45,"voornaam":"Liesbeth","tussenvoegsel":null,"achternaam":"Hermans","naam":"Liesbeth Hermans","email":"Liesbeth.hermans@live.nl","adres":"Turkoois 13","postcode":"5629 GN","woonplaats":"Eindhoven","telefoon":"06-20978825","gebdatum":"1946-07-28","lidsinds":null,"eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":46,"voornaam":"Linda","tussenvoegsel":"van der","achternaam":"Burg-Hamer","naam":"Linda van der Burg-Hamer","email":"lindahamer86@hotmail.com","adres":"De Tollande 52","postcode":"5528 BS","woonplaats":"Hoogeloon","telefoon":"06-30515032","gebdatum":"1986-06-11","lidsinds":"2025-06-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":47,"voornaam":"Lonneke ","tussenvoegsel":null,"achternaam":"Timmermans","naam":"Lonneke  Timmermans","email":"send2lon@hotmail.com","adres":"Bilderdijklaan2","postcode":"5531 TZ","woonplaats":"Bladel","telefoon":"06-50543068","gebdatum":"1974-07-19","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":"Secretaris","dagdelen":["wo 19.00 - 22.00"]},{"id":48,"voornaam":"Maarten","tussenvoegsel":"van","achternaam":"Dijk","naam":"Maarten van Dijk","email":"info@interpoint.nl","adres":"De Hoeve 7C","postcode":"5534 AC","woonplaats":"Netersel","telefoon":"06-22669837","gebdatum":"1975-07-27","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":"Bestuurslid","dagdelen":["wo 19.00 - 22.00"]},{"id":49,"voornaam":"Margaret","tussenvoegsel":null,"achternaam":"Janssen","naam":"Margaret Janssen","email":"margaretjanssen@hotmail.com","adres":"Marktstraat 18","postcode":"5531 AT","woonplaats":"Bladel","telefoon":"06-10997630","gebdatum":"1957-06-10","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 19.00 - 22.00"]},{"id":50,"voornaam":"Maria","tussenvoegsel":null,"achternaam":"Verweij-Bekx","naam":"Maria Verweij-Bekx","email":"h.verweij5@kpnmail.nl","adres":"Nachtgaallaan 6","postcode":"5561 TM","woonplaats":"Riethoven","telefoon":"06-23340607","gebdatum":"1951-03-03","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 19.00 - 22.00"]},{"id":51,"voornaam":"Marie-Hélène","tussenvoegsel":null,"achternaam":"Stokkink","naam":"Marie-Hélène Stokkink","email":"waterpainters@gmail.com","adres":"Quackelaer 37","postcode":"5521 BC","woonplaats":"Eersel","telefoon":"06-51941415","gebdatum":"1944-05-03","lidsinds":"2025-01-20","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":52,"voornaam":"Marietje","tussenvoegsel":"van de","achternaam":"Graaf","naam":"Marietje van de Graaf","email":" e.graaf3@upcmail.nl","adres":"Orion  16 b","postcode":"5527 CR","woonplaats":"Hapert","telefoon":"06-18915304","gebdatum":"1942-07-31","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":53,"voornaam":"Mariëtte","tussenvoegsel":"van","achternaam":"Sambeeck","naam":"Mariëtte van Sambeeck","email":"sambeeck16@hotmail.com","adres":"Heeleind 16a","postcode":"5531 CC","woonplaats":"Bladel","telefoon":"06-41547115","gebdatum":"1950-03-27","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":54,"voornaam":"Marijke","tussenvoegsel":null,"achternaam":"Donders","naam":"Marijke Donders","email":"marijdo@gmail.com","adres":"Doolandweg 34","postcode":"5531 PM","woonplaats":"Bladel","telefoon":"06-42842851","gebdatum":"1949-02-01","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":55,"voornaam":"Marjo","tussenvoegsel":null,"achternaam":"Vromans","naam":"Marjo Vromans","email":"marjovromans@gmail.com","adres":"Van Heinsbergdal 4","postcode":"5551 EZ","woonplaats":"Valkenswaard","telefoon":"06-57711110","gebdatum":"1959-06-28","lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":56,"voornaam":"Marjon","tussenvoegsel":"van der ","achternaam":"Mark","naam":"Marjon van der  Mark","email":"marjon@versid.nl","adres":"De Génestetlaan 2","postcode":"5531 TS","woonplaats":"Bladel","telefoon":"06-50847061","gebdatum":null,"lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":57,"voornaam":"Martin","tussenvoegsel":"van de ","achternaam":"Velden","naam":"Martin van de  Velden","email":"mgavdv@gmail.com","adres":"Keersop 2","postcode":"5551 TG","woonplaats":"Valkenswaard","telefoon":"06-49816585","gebdatum":null,"lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":58,"voornaam":"Marja","tussenvoegsel":null,"achternaam":"Cornelissen","naam":"Marja Cornelissen","email":"marja.cornelissen@hotmail.com","adres":"Emmaplein 11","postcode":"5531 HL","woonplaats":"Bladel","telefoon":"06-15193131","gebdatum":"1958-03-18","lidsinds":"2025-05-01","eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":59,"voornaam":"Marjoke","tussenvoegsel":"den","achternaam":"Engelsen","naam":"Marjoke den Engelsen","email":"mjde@upcmail.nl","adres":"Veilig Oord 72","postcode":"5531 XD","woonplaats":"Bladel","telefoon":"06-10869974","gebdatum":"1957-06-26","lidsinds":"2025-05-01","eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":60,"voornaam":"Mieke","tussenvoegsel":null,"achternaam":"Dirks","naam":"Mieke Dirks","email":"neptunus2005@live.nl","adres":"Neptunus 14","postcode":"5527 CD","woonplaats":"Hapert","telefoon":"06-53995221","gebdatum":"1949-10-26","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":61,"voornaam":"Monique","tussenvoegsel":null,"achternaam":"Kuijpers","naam":"Monique Kuijpers","email":"winkensm@planet.nl","adres":"Boshovensestraat 1A","postcode":"5561 AR","woonplaats":"Riethoven","telefoon":"06-10947545","gebdatum":"1957-10-10","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":62,"voornaam":"Nell","tussenvoegsel":null,"achternaam":"Beerends","naam":"Nell Beerends","email":"josnell21@hotmail.com","adres":"Van Dissellaan 21a","postcode":"5531 BP","woonplaats":"Bladel","telefoon":"06-50955241","gebdatum":"1955-05-09","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 19.00 - 22.00"]},{"id":63,"voornaam":"Nelya","tussenvoegsel":null,"achternaam":"Pelymskykh","naam":"Nelya Pelymskykh","email":"nelli_vaulina@hotmail.com","adres":"Herman Gorterlaan 5","postcode":"5531 SR","woonplaats":"Bladel","telefoon":"06-26464271","gebdatum":"1963-01-29","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 19.00 - 22.00"]},{"id":64,"voornaam":"Noor","tussenvoegsel":"van de","achternaam":"Molengraft","naam":"Noor van de Molengraft","email":"noor.vandemolengraft@icloud.com","adres":"Kempstraat 14","postcode":"5525 BD","woonplaats":"Duizel","telefoon":"06-82803213","gebdatum":"2008-07-19","lidsinds":"2026-05-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 19.00 - 22.00"]},{"id":65,"voornaam":"Noortje","tussenvoegsel":null,"achternaam":"Maas","naam":"Noortje Maas","email":"noormaas36@gmail.com","adres":"Van Rummenstraat 27","postcode":"5575  BT","woonplaats":"Luijksgestel","telefoon":"06-12238497","gebdatum":"1979-11-11","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":66,"voornaam":"Peter","tussenvoegsel":"van de","achternaam":"Wouw","naam":"Peter van de Wouw","email":"petervandewouw22@gmail.com","adres":"Dr. Cramerstraat 29","postcode":"5531 EP","woonplaats":"Bladel","telefoon":"06-20684111","gebdatum":"1953-04-25","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":67,"voornaam":"Petra","tussenvoegsel":null,"achternaam":"Tijsen","naam":"Petra Tijsen","email":"ptijsen19481@gmail.com","adres":"Fluitekruidbeemd 4","postcode":"5551 HT","woonplaats":"Valkenswaard","telefoon":"06-12480696","gebdatum":"1948-12-09","lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":68,"voornaam":"Renée","tussenvoegsel":null,"achternaam":"Holtzer","naam":"Renée Holtzer","email":"rholtzer@live.nl","adres":"Biestven 12","postcode":"5556 VR","woonplaats":"Valkenswaard","telefoon":"06-48833574","gebdatum":"1959-08-24","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":69,"voornaam":"Rieky","tussenvoegsel":"van","achternaam":"Iersel","naam":"Rieky van Iersel","email":"rieky52@gmail.com","adres":"Zwartakker 12","postcode":"5531 PB","woonplaats":"Bladel","telefoon":"06-15106477","gebdatum":"1952-05-25","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":70,"voornaam":"Rina","tussenvoegsel":null,"achternaam":"Kennis","naam":"Rina Kennis","email":"rinakennis@gmail.com","adres":"Aangelag 32","postcode":"5541 GJ","woonplaats":"Reusel","telefoon":"06-51891915","gebdatum":"1950-06-02","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]},{"id":71,"voornaam":"Sabien","tussenvoegsel":null,"achternaam":"Streppel","naam":"Sabien Streppel","email":"sabienstr@kpnmail.nl","adres":"Mozartstraat 63","postcode":"5481 LA","woonplaats":"Schijndel","telefoon":"06-30950982","gebdatum":null,"lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":72,"voornaam":"Sofiya","tussenvoegsel":null,"achternaam":"Verdonschot","naam":"Sofiya Verdonschot","email":"slonik@live.nl","adres":"Helleneind 16 A","postcode":"5531 BV ","woonplaats":"Bladel","telefoon":"06-19573732","gebdatum":"1987-07-18","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 19.00 - 22.00"]},{"id":73,"voornaam":"Thérèse","tussenvoegsel":null,"achternaam":"Pluijms","naam":"Thérèse Pluijms","email":"t.pluijms@hotmail.com","adres":"Jaques Perklaan 16","postcode":"5531 TR","woonplaats":"Bladel","telefoon":"06-12074150","gebdatum":"1978-12-10","lidsinds":"2025-04-01","eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":74,"voornaam":"Tinie","tussenvoegsel":null,"achternaam":"Gijbels","naam":"Tinie Gijbels","email":"t.gijbels@chello.nl","adres":"Kendersekker 10","postcode":"5541DL","woonplaats":"Reusel","telefoon":"0497-643216","gebdatum":null,"lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":75,"voornaam":"Thomas","tussenvoegsel":null,"achternaam":"Schilders","naam":"Thomas Schilders","email":"thomas-elly@famschilders.nl","adres":"Koekoeksbos 27","postcode":"5531 DZ","woonplaats":"Bladel","telefoon":"06-57246542","gebdatum":"1946-03-09","lidsinds":null,"eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":76,"voornaam":"Tonnie","tussenvoegsel":null,"achternaam":"Tholen","naam":"Tonnie Tholen","email":"tholentonnie@gmail.com ","adres":"Amalialaan 3","postcode":"5531 JJ","woonplaats":"Bladel","telefoon":"06-44150563","gebdatum":"1952-03-25","lidsinds":null,"eindelidmaat":"2025-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":77,"voornaam":"Tonny","tussenvoegsel":"van den","achternaam":"Tillaar","naam":"Tonny van den Tillaar","email":"advandentillaar@gmail.com","adres":"Nieuwstraat 46","postcode":"5527 AV","woonplaats":"Hapert","telefoon":"06-45661145","gebdatum":"1950-07-22","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 19.00 - 22.00"]},{"id":78,"voornaam":"Toos","tussenvoegsel":null,"achternaam":"Adams","naam":"Toos Adams","email":"toosadams@hotmail.com","adres":"Wilhelminalaan 88","postcode":"5531 HV","woonplaats":"Bladel","telefoon":"06-47890938","gebdatum":"1955-07-06","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["ma 11.00 - 16.00"]},{"id":79,"voornaam":"Truus","tussenvoegsel":null,"achternaam":"Stadler-Guns","naam":"Truus Stadler-Guns","email":"truus.stadler@kpnmail.nl","adres":"Midakkers 20","postcode":"5521 GM","woonplaats":"Eersel","telefoon":"06-27399706","gebdatum":"1948-02-06","lidsinds":"2025-01-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["wo 09.00 - 12.30"]},{"id":80,"voornaam":"Wilhelmien","tussenvoegsel":"van","achternaam":"Dijk","naam":"Wilhelmien van Dijk","email":"wihelmienvandijk@hotmail.com","adres":"Dokter Rauppstraat 69","postcode":"5571 CD","woonplaats":"Bergeijk","telefoon":"06-19667310","gebdatum":null,"lidsinds":null,"eindelidmaat":"2024-12-31","status":"inactief","functie":null,"dagdelen":[]},{"id":81,"voornaam":"Twan","tussenvoegsel":null,"achternaam":"Duis","naam":"Twan Duis","email":"t.duis@e-genius.nl","adres":"Lodewijk v Deijssellaan 19","postcode":"5531 SL","woonplaats":"Bladel","telefoon":"06-46346632","gebdatum":"1969-11-03","lidsinds":null,"eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 19.00 - 22.00"]},{"id":82,"voornaam":"Walter","tussenvoegsel":null,"achternaam":"Hendrickx","naam":"Walter Hendrickx","email":"walterhendrickx@hotmail.com","adres":"Veilig Oord 77","postcode":"5531 XD","woonplaats":"Bladel","telefoon":"06-30134889","gebdatum":"1956-03-30","lidsinds":"2025-04-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["di 10.00 - 16.00"]},{"id":83,"voornaam":"Zlaty","tussenvoegsel":null,"achternaam":"Grem","naam":"Zlaty Grem","email":"z.grem@outlook.com","adres":"Marktstaete","postcode":"5531 BR","woonplaats":"Bladel","telefoon":"06-10356375","gebdatum":"1950-04-15","lidsinds":"2025-05-01","eindelidmaat":null,"status":"actief","functie":null,"dagdelen":["do 09.30 - 16.00"]}];
const SEED_WORKSHOPS = [{"id":1,"titel":"blok 1 woe mid groep 5-8","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":2,"titel":"blok 1 zat mid groep 5-8","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":3,"titel":"blok 2 woe mid groep 5-8","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":4,"titel":"blok 1 25-26 woe mid groep 5-8 om 15:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":5,"titel":"blok 1 25-26 woe mid groep 5-8 om 13:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":6,"titel":"blok 2 25-26 woe mid groep 5-8 om 13:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":7,"titel":"blok 3 25-26 woe mid groep 5-8 om 13:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":8,"titel":"blok 2 25-26 woe mid groep 5-8 om 15:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":9,"titel":"blok 3 25-26 woe mid groep 5-8 om 15:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":10,"titel":"Blok 4 25-26 woe mid groep 5-8 om 13:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":11,"titel":"blok 4 25-26 woe mid groep 5-8 om 15:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":12,"titel":"Blok 5 25-26 woe mid groep 5-8 om 13:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":13,"titel":"Blok 5 25-26 woe mid groep 5-8 om 15:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":14,"titel":"Blok 1 26-27 woe mid groep 5-8 om 15:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"},{"id":15,"titel":"Blok 1 26-27 woe mid groep 5-8 om 13:30","soort":"Jeugdatelier","type":"reeks","datums":[],"dagdeel":"","locatie":"","bedrag":null,"maxDeelnemers":null,"status":"afgerond"}];
const SEED_WORKSHOP_INSCHRIJVINGEN = [{"id":1,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Bobbi Reijnders","email":"fennieh@hotmail.com","email2":"robreijnders2013@gmail.com","telefoon":"06-11645964","telefoon2":"06-47484061","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-14","notitie":""},{"id":2,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Thomas Pennings","email":"evastravers@gmail.com","email2":"","telefoon":"06-16120588","telefoon2":"","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-13","notitie":""},{"id":3,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Lissa Aling","email":"hm@sait.nl","email2":"","telefoon":"06-43588092","telefoon2":"06-54730232","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-13","notitie":""},{"id":4,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Adele Smirnova","email":"queetok@gmail.com","email2":"","telefoon":"06-45802541","telefoon2":"06-45497713","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-19","notitie":""},{"id":5,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Yenthe Kokken","email":"jhceelen@yahoo.com","email2":"","telefoon":"06-28187205","telefoon2":"","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-19","notitie":""},{"id":6,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Vajen van Maanen","email":"carlivanmaanen@gmail.com","email2":"","telefoon":"06-52431554","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-19","notitie":""},{"id":7,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Tess Wieland","email":"sjctorres@gmail.com","email2":"","telefoon":"06-45016896","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-20","notitie":""},{"id":8,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Lois Adams","email":"familieadamns@hotmail.com","email2":"","telefoon":"06-38399790","telefoon2":"","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-23","notitie":""},{"id":9,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Floor Bruggenwirth","email":"lvgorp@hotmail.com","email2":"","telefoon":"06-48938468","telefoon2":"","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-22","notitie":""},{"id":10,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Tess van de Spijker","email":"y_soontiens@live.nl","email2":"","telefoon":"06-46144746","telefoon2":"06-13087240","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-26","notitie":""},{"id":11,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Vajèn van Limpt","email":"Ron32vl@live.nl","email2":"Marieke15a@hotmail.com","telefoon":"06-21244669","telefoon2":"06-55360025","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-27","notitie":""},{"id":12,"workshopId":1,"herkomst":"extern","lidId":null,"naam":"Jip van Dingenen","email":"yvonne_lenaers@hotmail.com","email2":"","telefoon":"06-46784611","telefoon2":"","leeftijd":9,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-03-28","notitie":""},{"id":13,"workshopId":2,"herkomst":"extern","lidId":null,"naam":"Tess Waarma","email":"agent_sharon@hotmail.com","email2":"","telefoon":"06-30163450","telefoon2":"","leeftijd":9,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-05-01","notitie":""},{"id":14,"workshopId":2,"herkomst":"extern","lidId":null,"naam":"Sophie Wouters","email":"jannekewouterz@gmail.com","email2":"","telefoon":"06-21703282","telefoon2":"","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-05-04","notitie":""},{"id":15,"workshopId":2,"herkomst":"extern","lidId":null,"naam":"Bo Peijs","email":"sophiepeijs@live.nl","email2":"","telefoon":"06-53350015","telefoon2":"","leeftijd":9,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-05-06","notitie":""},{"id":16,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Elin Heesterbeek","email":"sandravhoof@hotmail.com","email2":"","telefoon":"06-44303575","telefoon2":"","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-05-08","notitie":""},{"id":17,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Eva-Marie Bampton","email":"carolinebampton@hotmail.com","email2":"","telefoon":"06-13519729","telefoon2":"","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-05-08","notitie":""},{"id":18,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Lois Adams","email":"familieadamns@hotmail.com","email2":"","telefoon":"06-38399790","telefoon2":"","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":19,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Tess Wieland","email":"sjctorres@gmail.com","email2":"","telefoon":"06-45016896","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":20,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Tess van de Spijker","email":"y_soontiens@live.nl","email2":"","telefoon":"06-46144746","telefoon2":"06-13087240","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":21,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Vajen van Maanen","email":"carlivanmaanen@gmail.com","email2":"","telefoon":"06-52431554","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":22,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Bobbi Reijnders","email":"fennieh@hotmail.com","email2":"robreijnders2013@gmail.com","telefoon":"06-11645964","telefoon2":"06-47484061","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":23,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Thomas Pennings","email":"evastravers@gmail.com","email2":"","telefoon":"06-16120588","telefoon2":"","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":24,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Lissa Aling","email":"hm@sait.nl","email2":"","telefoon":"06-43588092","telefoon2":"06-54730232","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":25,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Yenthe Kokken","email":"jhceelen@yahoo.com","email2":"","telefoon":"06-28187205","telefoon2":"","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":26,"workshopId":3,"herkomst":"extern","lidId":null,"naam":"Floor Bruggenwirth","email":"lvgorp@hotmail.com","email2":"","telefoon":"06-48938468","telefoon2":"","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":27,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Bo Peijs","email":"sophiepeijs@live.nl","email2":"","telefoon":"06-53350015","telefoon2":"","leeftijd":9,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":28,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Elin Heesterbeek","email":"sandravhoof@hotmail.com","email2":"","telefoon":"06-44303575","telefoon2":"","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":29,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Eva-Marie Bampton","email":"carolinebampton@hotmail.com","email2":"","telefoon":"06-13519729","telefoon2":"","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":30,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Fauve Taalman","email":"merelbeen@hotmail.com","email2":"","telefoon":"06-26330024","telefoon2":"06-48119739","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-26","notitie":""},{"id":31,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Floor Bruggenwirth","email":"lvgorp@hotmail.com","email2":"","telefoon":"06-48938468","telefoon2":"","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":32,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Lena van Boxtel","email":"vanboxtel.rianne@gmail.com","email2":"","telefoon":"06-43063157","telefoon2":"","leeftijd":12,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":33,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Lisa Visser","email":"kjcven@hotmail.com","email2":"","telefoon":"06-23871563","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-25","notitie":""},{"id":34,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Lois Adams","email":"familieadamns@hotmail.com","email2":"","telefoon":"06-38399790","telefoon2":"","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":35,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Sophie Wouters","email":"jannekewouterz@gmail.com","email2":"","telefoon":"06-21703282","telefoon2":"","leeftijd":11,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":36,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Tess van de Spijker","email":"y_soontiens@live.nl","email2":"","telefoon":"06-46144746","telefoon2":"06-13087240","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":37,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Vajen van Maanen","email":"carlivanmaanen@gmail.com","email2":"","telefoon":"06-52431554","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":38,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Yenthe Kokken","email":"jhceelen@yahoo.com","email2":"","telefoon":"06-28187205","telefoon2":"","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":39,"workshopId":4,"herkomst":"extern","lidId":null,"naam":"Zoë van der Zanden","email":"wendybloks@hotmail.com","email2":"","telefoon":"06-25055518","telefoon2":"","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-07-22","notitie":""},{"id":40,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Coco Prinsen","email":"brittheeren@hotmail.com","email2":"","telefoon":"06-18162596","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-23","notitie":""},{"id":41,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Evi Gooskens","email":"gooskenssanne@gmail.com","email2":"","telefoon":"06-50222809","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-19","notitie":""},{"id":42,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Fee van Huijkelom","email":"edomarjon@hotmail.com","email2":"","telefoon":"06-27072680","telefoon2":"06-13724277","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-20","notitie":""},{"id":43,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Fenne Fabrie","email":"LindaHuybregts@hotmail.com","email2":"","telefoon":"06-13543546","telefoon2":"","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-22","notitie":""},{"id":44,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Fienne Bax","email":"s_geraerts@hotmail.com","email2":"","telefoon":"06-18238291","telefoon2":"","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-20","notitie":""},{"id":45,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Izebeau te Riele","email":"anitahabraken_@hotmail.com","email2":"","telefoon":"06-13183642","telefoon2":"","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-24","notitie":""},{"id":46,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Jinze van Limpt","email":"rowey.schoofs@hotmail.nl","email2":"","telefoon":"06-22855123","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-24","notitie":""},{"id":47,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Kees Brekelmans","email":"s_phenninckx@hotmail.com","email2":"","telefoon":"06-30717589","telefoon2":"","leeftijd":9,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-19","notitie":""},{"id":48,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Laure van den Borne","email":"leonievanzullichem@gmail.com","email2":"","telefoon":"06-51941779","telefoon2":"06-51553014","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-21","notitie":""},{"id":49,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Lola Prinsen","email":"brittheeren@hotmail.com","email2":"","telefoon":"06-18162596","telefoon2":"","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-23","notitie":""},{"id":50,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Marit Jansen","email":"ilseverkley@hotmaio.com","email2":"","telefoon":"06-17666625","telefoon2":"06-17221155","leeftijd":7,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"2025-06-23","notitie":""},{"id":51,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Sanna van den Borne","email":"leonievanzullichem@gmail.com","email2":"roelvandenborne@hotmail.com","telefoon":"06-51941779","telefoon2":"06-51553014","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":52,"workshopId":5,"herkomst":"extern","lidId":null,"naam":"Tess Waarma","email":"agent_sharon@hotmail.com","email2":"","telefoon":"06-30163450","telefoon2":"","leeftijd":9,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":53,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Coco Prinsen","email":"brittheeren@hotmail.com","email2":"","telefoon":"06-18162596","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":54,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Evi Gooskens","email":"gooskenssanne@gmail.com","email2":"","telefoon":"06-50222809","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":55,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Fee van Huijkelom","email":"edomarjon@hotmail.com","email2":"","telefoon":"06-27072680","telefoon2":"06-13724277","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":56,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Fienne Bax","email":"s_geraerts@hotmail.com","email2":"","telefoon":"06-18238291","telefoon2":"","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":57,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Izebeau te Riele","email":"anitahabraken_@hotmail.com","email2":"","telefoon":"06-13183642","telefoon2":"","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":58,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Jinze van Limpt","email":"rowey.schoofs@hotmail.nl","email2":"","telefoon":"06-22855123","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":59,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Lola Prinsen","email":"brittheeren@hotmail.com","email2":"","telefoon":"06-18162596","telefoon2":"","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":60,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Lynn Vernie","email":"dorindavernie@hotmail.com","email2":"","telefoon":"06-13616019","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":61,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Mara van Straaten","email":"timmersmansm19@gmail.com","email2":"","telefoon":"06-30265431","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":62,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Marit Jansen","email":"ilseverkley@hotmaio.com","email2":"","telefoon":"06-17666625","telefoon2":"06-17221155","leeftijd":7,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":63,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Sanna van den Borne","email":"leonievanzullichem@gmail.com","email2":"roelvandenborne@hotmail.com","telefoon":"06-51941779","telefoon2":"06-51553014","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":64,"workshopId":6,"herkomst":"extern","lidId":null,"naam":"Tess Waarma","email":"agent_sharon@hotmail.com","email2":"","telefoon":"06-30163450","telefoon2":"","leeftijd":9,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":65,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Bo Peijs","email":"sophiepeijs@live.nl","email2":"","telefoon":"06-53350015","telefoon2":"","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":66,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Coco Prinsen","email":"brittheeren@hotmail.com","email2":"","telefoon":"06-18162596","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":67,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Eva van Deursen","email":"natasjavandeursen@gmail.com","email2":"","telefoon":"06-53284847","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":68,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Evi Gooskens","email":"gooskenssanne@gmail.com","email2":"","telefoon":"06-50222809","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":69,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Izebeau te Riele","email":"anitahabraken_@hotmail.com","email2":"","telefoon":"06-13183642","telefoon2":"","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":70,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Jinze van Limpt","email":"rowey.schoofs@hotmail.nl","email2":"","telefoon":"06-22855123","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":71,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Lola Prinsen","email":"brittheeren@hotmail.com","email2":"","telefoon":"06-18162596","telefoon2":"","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":72,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Lynn Vernie","email":"dorindavernie@hotmail.com","email2":"","telefoon":"06-13616019","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":73,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Mara van Straaten","email":"timmersmansm19@gmail.com","email2":"","telefoon":"06-30265431","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":74,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Marit Jansen","email":"ilseverkley@hotmaio.com","email2":"","telefoon":"06-17666625","telefoon2":"06-17221155","leeftijd":7,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":75,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Sanna van den Borne","email":"leonievanzullichem@gmail.com","email2":"roelvandenborne@hotmail.com","telefoon":"06-51941779","telefoon2":"06-51553014","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":76,"workshopId":7,"herkomst":"extern","lidId":null,"naam":"Tess Waarma","email":"agent_sharon@hotmail.com","email2":"","telefoon":"06-30163450","telefoon2":"","leeftijd":9,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":77,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Evy Hendriks","email":"familie@tobbert.nl","email2":"","telefoon":"06-51538911","telefoon2":"","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":78,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Fauve Taalman","email":"merelbeen@hotmail.com","email2":"","telefoon":"06-26330024","telefoon2":"06-48119739","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":79,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Ize Sebregts","email":"annetaalders@hotmail.com","email2":"","telefoon":"06-44308402","telefoon2":"06-21945176","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":80,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Jackie Kavelaars","email":"a.kavelaars@hotmail.com","email2":"","telefoon":"06-11079126","telefoon2":"06-15469540","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":81,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Lisa Visser","email":"kjcven@hotmail.com","email2":"","telefoon":"06-23871563","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":82,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Livia Rombouts","email":"juliadurczok@gmail.com","email2":"","telefoon":"06-33862192","telefoon2":"","leeftijd":9,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":83,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Roqi Snijders","email":"smg.nijhuis@gmail.com","email2":"marcelsnijders77@gmail.com","telefoon":"06-24661070","telefoon2":"06-46392188","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":"4 lessen","notitie":""},{"id":84,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Tess van de Spijker","email":"y_soontiens@live.nl","email2":"","telefoon":"06-46144746","telefoon2":"06-13087240","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":85,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Vera Tenbült ","email":"kvdheijden83@gmail.com","email2":"","telefoon":"06-24428404","telefoon2":"06-55392857","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":86,"workshopId":8,"herkomst":"extern","lidId":null,"naam":"Zoë van der Zanden","email":"wendybloks@hotmail.com","email2":"","telefoon":"06-25055518","telefoon2":"","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":87,"workshopId":9,"herkomst":"extern","lidId":null,"naam":"Emilie van Gestel","email":"carlijnegoossens@hotmail.com","email2":"","telefoon":"06-52157866","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":88,"workshopId":9,"herkomst":"extern","lidId":null,"naam":"Evi Hoeks","email":"addie_lauwers@hotmail.com","email2":"","telefoon":"06-46023542","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":89,"workshopId":9,"herkomst":"extern","lidId":null,"naam":"Ize Sebregts","email":"annetaalders@hotmail.com","email2":"","telefoon":"06-44308402","telefoon2":"06-21945176","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":90,"workshopId":9,"herkomst":"extern","lidId":null,"naam":"Livia Rombouts","email":"juliadurczok@gmail.com","email2":"","telefoon":"06-33862192","telefoon2":"","leeftijd":9,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":91,"workshopId":9,"herkomst":"extern","lidId":null,"naam":"Sophie Merks","email":"judith.theuws@hotmail.com","email2":"","telefoon":"06-51955425","telefoon2":"","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":92,"workshopId":9,"herkomst":"extern","lidId":null,"naam":"Tess van de Spijker","email":"y_soontiens@live.nl","email2":"","telefoon":"06-46144746","telefoon2":"06-13087240","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":93,"workshopId":9,"herkomst":"extern","lidId":null,"naam":"Tess Wouters","email":"susanvandenhout@gmail.com","email2":"","telefoon":"06-10200655","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":94,"workshopId":9,"herkomst":"extern","lidId":null,"naam":"Vera Tenbült ","email":"kvdheijden83@gmail.com","email2":"","telefoon":"06-24428404","telefoon2":"06-55392857","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":95,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Bo Peijs","email":"sophiepeijs@live.nl","email2":"","telefoon":"06-53350015","telefoon2":"","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":96,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Coco Prinsen","email":"brittheeren@hotmail.com","email2":"","telefoon":"06-18162596","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":97,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Eva van Deursen","email":"natasjavandeursen@gmail.com","email2":"","telefoon":"06-53284847","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":98,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Evi Gooskens","email":"gooskenssanne@gmail.com","email2":"","telefoon":"06-50222809","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":99,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Izebeau te Riele","email":"anitahabraken_@hotmail.com","email2":"","telefoon":"06-13183642","telefoon2":"","leeftijd":10,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":100,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Jinze van Limpt","email":"rowey.schoofs@hotmail.nl","email2":"","telefoon":"06-22855123","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":101,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Lola Prinsen","email":"brittheeren@hotmail.com","email2":"","telefoon":"06-18162596","telefoon2":"","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":102,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Lynn Vernie","email":"dorindavernie@hotmail.com","email2":"","telefoon":"06-13616019","telefoon2":"","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":103,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Marit Jansen","email":"ilseverkley@hotmaio.com","email2":"","telefoon":"06-17666625","telefoon2":"06-17221155","leeftijd":7,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":104,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Sanna van den Borne","email":"leonievanzullichem@gmail.com","email2":"roelvandenborne@hotmail.com","telefoon":"06-51941779","telefoon2":"06-51553014","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":105,"workshopId":10,"herkomst":"extern","lidId":null,"naam":"Tess Waarma","email":"agent_sharon@hotmail.com","email2":"","telefoon":"06-30163450","telefoon2":"","leeftijd":9,"groep":5,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":106,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Emilie van Gestel","email":"carlijnegoossens@hotmail.com","email2":"","telefoon":"06-52157866","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":107,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Evi Hoeks","email":"addie_lauwers@hotmail.com","email2":"","telefoon":"06-46023542","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":108,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Ize Sebregts","email":"annetaalders@hotmail.com","email2":"","telefoon":"06-44308402","telefoon2":"06-21945176","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":109,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Livia Rombouts","email":"juliadurczok@gmail.com","email2":"","telefoon":"06-33862192","telefoon2":"","leeftijd":9,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":110,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Sophie Merks","email":"judith.theuws@hotmail.com","email2":"","telefoon":"06-51955425","telefoon2":"","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":111,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Tess van de Spijker","email":"y_soontiens@live.nl","email2":"","telefoon":"06-46144746","telefoon2":"06-13087240","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":112,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Tess Wouters","email":"susanvandenhout@gmail.com","email2":"","telefoon":"06-10200655","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":113,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Vera Tenbült ","email":"kvdheijden83@gmail.com","email2":"","telefoon":"06-24428404","telefoon2":"06-55392857","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":114,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Vivien Strijbos","email":"ploegmakerslinda@hotmail.com","email2":"","telefoon":"06-22917656","telefoon2":"","leeftijd":10,"groep":6,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":115,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Noor van Kreij","email":"carlievandenborne@hotmail.com","email2":"","telefoon":"06-25477463","telefoon2":"06-27567636","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":116,"workshopId":11,"herkomst":"extern","lidId":null,"naam":"Amy van Vessem","email":"daisyvanvessem@hotmail.com","email2":"","telefoon":"06-13110295","telefoon2":"","leeftijd":11,"groep":8,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":117,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Ivy Hurkmans","email":"brittsijtsma@gmail.com","email2":"","telefoon":"06-41065556","telefoon2":"","leeftijd":8,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":118,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Tess Waarma","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":119,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Evi Gooskens","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":120,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Sanna van den Borne","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":false,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":121,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Marit Jansen","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":122,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Coco Prinsen","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":123,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Jinze van Limpt","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":124,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Lynn Vernie","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":125,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Eva van Deursen","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":126,"workshopId":12,"herkomst":"extern","lidId":null,"naam":"Nune van der Heijden","email":"robinvanherpt@gmail.com","email2":"","telefoon":"06-22864053","telefoon2":"","leeftijd":7,"groep":4,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":127,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Tess van de Spijker","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":128,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Ize Sebregts","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":129,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Livia Rombouts","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":130,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Vera Tenbült ","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":131,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Sophie Merks","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":132,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Emilie van Gestel","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":133,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Vivien Strijbos","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":134,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Noor van Kreij","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":false,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":135,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Amy van Vessem","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":136,"workshopId":13,"herkomst":"extern","lidId":null,"naam":"Mick Lepelaars","email":"","email2":"","telefoon":"","telefoon2":"","leeftijd":null,"groep":null,"status":"ingeschreven","betaald":true,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":137,"workshopId":14,"herkomst":"extern","lidId":null,"naam":"Fenna Koolen","email":"spj.tops@gmail.com","email2":"","telefoon":"06-14617112","telefoon2":"06-46109970","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":false,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":138,"workshopId":15,"herkomst":"extern","lidId":null,"naam":"Saar van der Kroon","email":"elze_maas@hotmail.com","email2":"Martijnvdkroon@hotmail.com","telefoon":"06-27507009","telefoon2":"06-41682796","leeftijd":10,"groep":7,"status":"ingeschreven","betaald":false,"bedrag":null,"datumInschrijving":null,"notitie":""},{"id":139,"workshopId":15,"herkomst":"extern","lidId":null,"naam":"Pien van der Kroon","email":"elze_maas@hotmail.com","email2":"Martijnvdkroon@hotmail.com","telefoon":"06-27507009","telefoon2":"06-41682796","leeftijd":8,"groep":5,"status":"ingeschreven","betaald":false,"bedrag":null,"datumInschrijving":null,"notitie":""}];
const SEED_WORKSHOP_SOORTEN = ["Jeugdatelier","Modeltekenen","Werken met paletmes","Portretschilderen","Schilderen"];
const SEED_TX = [{"id":1,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-01","grootboek_code":"1130","grootboek_naam":"Rabobank .319","bedrag":139.23,"omschrijving":"overboeking rente spaarrekening"},{"id":2,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-01","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":42,"omschrijving":"v Sambeek van Rijswijk afspraak Dorien"},{"id":3,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-01","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Tonnie Tholen"},{"id":4,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-02","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Gerthy Fleskens"},{"id":5,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-02","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Angelina van Kemenade"},{"id":6,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-02","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-17.16,"omschrijving":"Kosten"},{"id":7,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-03","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Truus Stadler"},{"id":8,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-03","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Noortje Maas"},{"id":9,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-03","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Toos Adams"},{"id":10,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-06","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Dick Scheepmaker"},{"id":11,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-06","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Jeanne van Leeuwen"},{"id":12,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-08","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":" Karien vd Heijden"},{"id":13,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-08","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Margaret Janssen"},{"id":14,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-12","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Huub vd Krogt"},{"id":15,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-13","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Hanneke de Groot"},{"id":16,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-14","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Peter vd Wouw"},{"id":17,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-15","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Dorien Reijnders"},{"id":18,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-15","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Twan Duis"},{"id":19,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-15","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Hettie Daamen"},{"id":20,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-16","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Daan vd Enden"},{"id":21,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-16","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Maria Verweij-Bekx"},{"id":22,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-17","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Marijke Donders"},{"id":23,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-17","grootboek_code":"8015","grootboek_naam":"Subsidies en bijdragen","bedrag":1500,"omschrijving":"voorschot subsidie gemeente Bladel"},{"id":24,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-17","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Monique Kuijpers"},{"id":25,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-18","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Anneke Lablans"},{"id":26,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-18","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Hennie Walenberg"},{"id":27,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-19","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Liesbeth Adams"},{"id":28,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-20","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Anton Cools"},{"id":29,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-20","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Guus Goossens"},{"id":30,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-20","grootboek_code":"8021","grootboek_naam":"Spaaracties overig","bedrag":267,"omschrijving":"Spek je spaarkas actie Sniederspassage"},{"id":31,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-22","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Ingrid de Wit"},{"id":32,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-22","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Liesbeth Hermans"},{"id":33,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-22","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Marietje de Graaf"},{"id":34,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-24","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Arianne van Hout"},{"id":35,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-24","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Henny Herps"},{"id":36,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-24","grootboek_code":"4620","grootboek_naam":"Representatiekosten","bedrag":-30,"omschrijving":"Bloemen Dorien aan Toos Adams"},{"id":37,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-24","grootboek_code":"4300","grootboek_naam":"Huur","bedrag":-2616.94,"omschrijving":"Huur atelier 2024"},{"id":38,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-24","grootboek_code":"4310","grootboek_naam":"Energiekosten","bedrag":-1200.86,"omschrijving":"Energiekosten atelier 2024"},{"id":39,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-24","grootboek_code":"4310","grootboek_naam":"Energiekosten","bedrag":7.05,"omschrijving":"afrekening energiekosten 2023"},{"id":40,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-24","grootboek_code":"1500","grootboek_naam":"Betaalde BTW","bedrag":-250.7,"omschrijving":"21% over energiekosten en teruggave"},{"id":41,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-24","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Mieke Dirks"},{"id":42,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-27","grootboek_code":"4650","grootboek_naam":"Verzekeringen","bedrag":-152.21,"omschrijving":"Quintes Aansprakelijkheidsverz"},{"id":43,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-27","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Gijsbert Jansen"},{"id":44,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-29","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Tonny vd Tillaar"},{"id":45,"rekening":"908","jaar":2025,"maand":1,"datum":"2025-01-30","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Marie-Hélène Stokkink"},{"id":46,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-01-30","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Elly Verrijt"},{"id":47,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-03","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-30.12,"omschrijving":"kosten"},{"id":48,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Maarten van Dijk"},{"id":49,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-07","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Rina Kennis"},{"id":50,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-07","grootboek_code":"1050","grootboek_naam":"Kruisposten","bedrag":-94.83,"omschrijving":"betaling Infomedics"},{"id":51,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-08","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Ivo Strouken"},{"id":52,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-09","grootboek_code":"1050","grootboek_naam":"Kruisposten","bedrag":94.83,"omschrijving":"Terugbetaling Infomedics G Jansen"},{"id":53,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-09","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-31.06,"omschrijving":"Terugbetaling Cor van Gool"},{"id":54,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-10","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":31.06,"omschrijving":"Boodschappen Cor van Gool"},{"id":55,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-12","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Jeanette Graamans"},{"id":56,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-13","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Sofia Verdonschot"},{"id":57,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-13","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Nelya Pelymskykh"},{"id":58,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-18","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Brigitte Burgmans"},{"id":59,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-19","grootboek_code":"4320","grootboek_naam":"Inventaris atelier","bedrag":-513.16,"omschrijving":"Stellingkast.nl inrichting atelier"},{"id":60,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-19","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-106.74,"omschrijving":"inkoop nieuwjaarsborrel Lonneke Timmermans"},{"id":61,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-20","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Thomas Schilders"},{"id":62,"rekening":"908","jaar":2025,"maand":2,"datum":"2025-02-20","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":96.25,"omschrijving":"Andrea Boullart"},{"id":63,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-03","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-23.52,"omschrijving":"rente en bankkosten"},{"id":64,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-12","grootboek_code":"8015","grootboek_naam":"Subsidies en bijdragen","bedrag":12.5,"omschrijving":"Huur panelen Expo Kunstuitleen Veldhoven"},{"id":65,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-12","grootboek_code":"4620","grootboek_naam":"Representatiekosten","bedrag":-50,"omschrijving":"Interpoint Flyers en Poster Jeugdatelier"},{"id":66,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lissa Aling"},{"id":67,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Thomas Pennings"},{"id":68,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-14","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Bobbi Reijnders"},{"id":69,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-15","grootboek_code":"4390","grootboek_naam":"Overige huisvestingskosten","bedrag":-125,"omschrijving":"realisatie droogrek van Spreeuwel"},{"id":70,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-19","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Adelaida Smirnova"},{"id":71,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Yenthe Kokken"},{"id":72,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-22","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess vd Spijker"},{"id":73,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-22","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Floor Bruggenwirth"},{"id":74,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vajen van Limpt"},{"id":75,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-27","grootboek_code":"4320","grootboek_naam":"Inventaris atelier","bedrag":-76.9,"omschrijving":"sleutelkastje Kabelshop.nl"},{"id":76,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-28","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Jip van Dingenen"},{"id":77,"rekening":"908","jaar":2025,"maand":3,"datum":"2025-03-28","grootboek_code":"4005","grootboek_naam":"Schoonmaakkosten atelier","bedrag":-112.05,"omschrijving":"Papierrollen + dispenser"},{"id":78,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-01","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vajen van Maanen"},{"id":79,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-01","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lois Adams"},{"id":80,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-01","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Wieland"},{"id":81,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-23.52,"omschrijving":"rente en bankkosten"},{"id":82,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-02","grootboek_code":"4390","grootboek_naam":"Overige huisvestingskosten","bedrag":-24.75,"omschrijving":"Scharniertjes tafelezelw Cor  v G"},{"id":83,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-02","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-49.17,"omschrijving":"matriaal jeugdatelier Daan vd E"},{"id":84,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-07","grootboek_code":"8015","grootboek_naam":"Subsidies en bijdragen","bedrag":1332.77,"omschrijving":"Restant saldo de Rietpen"},{"id":85,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-07","grootboek_code":"4005","grootboek_naam":"Schoonmaakkosten atelier","bedrag":-4.29,"omschrijving":"zachte zeep Toos Adams"},{"id":86,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-08","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-10.65,"omschrijving":"materiaal jeugdatelier Angelina v K"},{"id":87,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-08","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-195.93,"omschrijving":"materiaal jeugdatelier Daan vd E"},{"id":88,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-09","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-20.58,"omschrijving":"Koffie-thee Dorien R"},{"id":89,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-19","grootboek_code":"4005","grootboek_naam":"Schoonmaakkosten atelier","bedrag":-13,"omschrijving":"betaling stort veldhoven Daan vd E"},{"id":90,"rekening":"908","jaar":2025,"maand":4,"datum":"2025-04-25","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":78.75,"omschrijving":"Therese Vrielink"},{"id":91,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-01","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Waarma"},{"id":92,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-01","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-21.6,"omschrijving":"drankjes bestuur"},{"id":93,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-01","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-150,"omschrijving":"vergoeding docent blok 1 Daan vd E"},{"id":94,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-02","grootboek_code":"4390","grootboek_naam":"Overige huisvestingskosten","bedrag":-97,"omschrijving":"WOZ gemeente werf gemeente Bladel"},{"id":95,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-02","grootboek_code":"4390","grootboek_naam":"Overige huisvestingskosten","bedrag":-37.96,"omschrijving":"meubeltransporters Gijsbert J"},{"id":96,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-02","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-22.95,"omschrijving":"rente en bankkosten"},{"id":97,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-04","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sophie Wouters"},{"id":98,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-05","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lois Adams"},{"id":99,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-07","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Martha Obbink"},{"id":100,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-07","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Bo Peijs"},{"id":101,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-08","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-150,"omschrijving":"vergoeding docent blok 1 Angelina v K"},{"id":102,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-09","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Eva-Marie Bampton"},{"id":103,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Wieland"},{"id":104,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-12","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Elin Heesterbeek"},{"id":105,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-14","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess vd Spijker"},{"id":106,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-20","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-67.68,"omschrijving":"materiaal jeugdatelier Daan vd E"},{"id":107,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-20","grootboek_code":"4440","grootboek_naam":"Algemene ledenvergadering","bedrag":-177.5,"omschrijving":"ALV jan 25 BRUIS"},{"id":108,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-20","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-211.5,"omschrijving":"BRUIS april 25"},{"id":109,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-21","grootboek_code":"4611","grootboek_naam":"Website","bedrag":-75,"omschrijving":"AVG pakket graaggoedonline.nl"},{"id":110,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-21","grootboek_code":"4611","grootboek_naam":"Website","bedrag":-15.75,"omschrijving":"BTW tbv AVG pakket graaggoedonline.nl"},{"id":111,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-22","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vajen van Maanen"},{"id":112,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-23","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-37.72,"omschrijving":"materiaal jeugdatelier Angelina v K"},{"id":113,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-23","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-11.3,"omschrijving":"materiaal jeugdatelier Daan vd E"},{"id":114,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-25","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-70,"omschrijving":"materiaal jeugdatelier Daan vd E"},{"id":115,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Bobbi Reijnders"},{"id":116,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Floor Bruggenwirth"},{"id":117,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Thomas Pennings"},{"id":118,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Yenthe Kokken"},{"id":119,"rekening":"908","jaar":2025,"maand":5,"datum":"2025-05-28","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lissa Aling"},{"id":120,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-02","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Zlaty Grem"},{"id":121,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-02","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-29.61,"omschrijving":"rente en bankkosten"},{"id":122,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-03","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-25,"omschrijving":"BRUIS mei 25-1"},{"id":123,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-04","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":-5,"omschrijving":"restitutie 1 lesdag Tess Waarma"},{"id":124,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-04","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":-5,"omschrijving":"restitutie 1 lesdag Vajen van Maanen"},{"id":125,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":-35,"omschrijving":"restitutie Zlaty Grem teveel betaald"},{"id":126,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-04","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-137.5,"omschrijving":"BRUIS mei-2"},{"id":127,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-12","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-210,"omschrijving":"vergoeding docent blok 2 Daan vd E"},{"id":128,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-12","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-240,"omschrijving":"vergoeding docent blok 2 Angelina v K"},{"id":129,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-13","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":61.25,"omschrijving":"Jolanda Spaan"},{"id":130,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-17","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":-96.25,"omschrijving":"Restitutie Martha Obbink"},{"id":131,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-19","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Kees Brekelmans"},{"id":132,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-19","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Evi Gooskens"},{"id":133,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-20","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":61.25,"omschrijving":"Linda vd Berg"},{"id":134,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-20","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Fiene Bax"},{"id":135,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sanne van den Borne"},{"id":136,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Laure van den Borne"},{"id":137,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-23","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Marit Janssen"},{"id":138,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-24","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Jinze van Limpt"},{"id":139,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-24","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Izebeau te Riele"},{"id":140,"rekening":"908","jaar":2025,"maand":6,"datum":"2025-06-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess vd Spijker"},{"id":141,"rekening":"908","jaar":2025,"maand":7,"datum":"2025-07-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-24.84,"omschrijving":"rente en bankkosten"},{"id":142,"rekening":"908","jaar":2025,"maand":7,"datum":"2025-07-08","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-43.4,"omschrijving":"BRUIS juni-1"},{"id":143,"rekening":"908","jaar":2025,"maand":7,"datum":"2025-07-08","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-127.75,"omschrijving":"BRUIS juni-2"},{"id":144,"rekening":"908","jaar":2025,"maand":7,"datum":"2025-07-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sophie Wouters"},{"id":145,"rekening":"908","jaar":2025,"maand":7,"datum":"2025-07-16","grootboek_code":"4320","grootboek_naam":"Inventaris atelier","bedrag":-96.99,"omschrijving":"VidaXL kastje atelier"},{"id":146,"rekening":"908","jaar":2025,"maand":7,"datum":"2025-07-22","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Zoë vd Zanden"},{"id":147,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-21.39,"omschrijving":"rente en bankkosten"},{"id":148,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-11","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-143.63,"omschrijving":"materiaal jeugdatelier Daan vd E"},{"id":149,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-14","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":70,"omschrijving":"Marjoke den Engelsen"},{"id":150,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-14","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Fee van Huijkelom"},{"id":151,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-17","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":43.75,"omschrijving":"Joyce vd Kam"},{"id":152,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-19","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-36.63,"omschrijving":"materiaal jeugdatelier Daan vd E"},{"id":153,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-20","grootboek_code":"4390","grootboek_naam":"Overige huisvestingskosten","bedrag":-5.49,"omschrijving":"baterijen sleutelkastje"},{"id":154,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-22","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lisa Visser"},{"id":155,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-24","grootboek_code":"4440","grootboek_naam":"Algemene ledenvergadering","bedrag":-60,"omschrijving":"eerste betaling ruimte nieuwjaarsborrel"},{"id":156,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-25","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lena v Boxtel"},{"id":157,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-25","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":70,"omschrijving":"Marja Cornelissen"},{"id":158,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Yenthe Kokken"},{"id":159,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Elin Heesterbeek"},{"id":160,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Eva-Marie Bampton"},{"id":161,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-27","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-60,"omschrijving":"vergoeding docent zomerworkshops Daan vd E"},{"id":162,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-29","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Fenne Fabrie"},{"id":163,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-29","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Waarma"},{"id":164,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-29","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Fauve Taalman"},{"id":165,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-31","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Coco Prinsen"},{"id":166,"rekening":"908","jaar":2025,"maand":8,"datum":"2025-08-31","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lola Prinsen"},{"id":167,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-24.69,"omschrijving":"rente en bankkosten"},{"id":168,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-01","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Bo Peijs"},{"id":169,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-01","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lois Adams"},{"id":170,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-02","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vajen van Maanen"},{"id":171,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-08","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Floor Bruggenwirth"},{"id":172,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-11","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-15,"omschrijving":"BRUIS 2025 0318 aug-2"},{"id":173,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-11","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-188.5,"omschrijving":"BRUIS 2025 0327 juli-aug"},{"id":174,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-11","grootboek_code":"4440","grootboek_naam":"Algemene ledenvergadering","bedrag":-257.2,"omschrijving":"BRUIS 2025 0320 sep-1 ALV"},{"id":175,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-11","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-16.99,"omschrijving":"mat. Jeugd Angelina vK"},{"id":176,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-11","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-3.96,"omschrijving":"mat. Jeugd Angelina vK"},{"id":177,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-15","grootboek_code":"4620","grootboek_naam":"Representatiekosten","bedrag":-29.45,"omschrijving":"Bloemen afscheid Toos A door Cor vG"},{"id":178,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-17","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-6.98,"omschrijving":"Jeugd  siroop  Daan vdE"},{"id":179,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-17","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-140.48,"omschrijving":"mat. Jeugd sep Daan  vdE"},{"id":180,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-17","grootboek_code":"4320","grootboek_naam":"Inventaris atelier","bedrag":-55.8,"omschrijving":"opbergboxen 2x GijsbertJ"},{"id":181,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-17","grootboek_code":"4320","grootboek_naam":"Inventaris atelier","bedrag":-553.51,"omschrijving":"uitbreiding stellingkast atelier"},{"id":182,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-17","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-112.78,"omschrijving":"diverse bonnen LonnekeT"},{"id":183,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-17","grootboek_code":"4640","grootboek_naam":"Bestuurskosten","bedrag":-410.65,"omschrijving":"etentje bestuur afscheid Toos"},{"id":184,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-18","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":-43.75,"omschrijving":"restitutie contributie Joyce vd Kam-Wilderbeek"},{"id":185,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-19","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-17.68,"omschrijving":"mat. Jeugd Daan vdE"},{"id":186,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-26","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Fiene Bax"},{"id":187,"rekening":"908","jaar":2025,"maand":9,"datum":"2025-09-26","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess vd Spijker"},{"id":188,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-01","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-165,"omschrijving":"BRUIS 20250369 sep-3"},{"id":189,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-18.54,"omschrijving":"rente en bankkosten"},{"id":190,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-03","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-270,"omschrijving":"vergoeding docent blok 1 sept Daan vdE"},{"id":191,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-03","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-270,"omschrijving":"vergoeding docent blok1 sept Angelina vK"},{"id":192,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-06","grootboek_code":"4611","grootboek_naam":"Website","bedrag":-421.08,"omschrijving":"GraagGoedOnline.nl"},{"id":193,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-06","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Fee van Huijkelom"},{"id":194,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-06","grootboek_code":"1130","grootboek_naam":"Rabobank .319","bedrag":1500,"omschrijving":"ophoging saldo"},{"id":195,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-07","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":26.25,"omschrijving":"Elly Hovens"},{"id":196,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-08","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lynn Vernie"},{"id":197,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-08","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-62.54,"omschrijving":"Posters etc Expositie 2025 Interpoint"},{"id":198,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-09","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":26.25,"omschrijving":"Karin Tonneijk"},{"id":199,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-14","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Evy Hendriks"},{"id":200,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-19","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Mara van Straaten"},{"id":201,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-20","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vera Tenbult"},{"id":202,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-20","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-44,"omschrijving":"BRUIS 20250387 okt-2"},{"id":203,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-20","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-5,"omschrijving":"BRUIS 20250373 okt-1"},{"id":204,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-20","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Jackie Kavelaars"},{"id":205,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-20","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Ize Sebregts"},{"id":206,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Jinze van Limpt"},{"id":207,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Waarma"},{"id":208,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lisa Visser"},{"id":209,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Marit Janssen"},{"id":210,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Evi Gooskens"},{"id":211,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Izebeau te Riele"},{"id":212,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Livia Rombouts"},{"id":213,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Fauve Taalman"},{"id":214,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-22","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-10,"omschrijving":"Cor vG Visit Bladel Bord expositie"},{"id":215,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-22","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Zoë vd Zanden"},{"id":216,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-22","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":50,"omschrijving":"Lola en Coco Prinsen"},{"id":217,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-22","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-40.17,"omschrijving":"mat jeugd Angelia vK"},{"id":218,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-22","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sanne van den Borne"},{"id":219,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-24","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":20,"omschrijving":"Roqi Nijhuis 1 les minder"},{"id":220,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-26","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-44,"omschrijving":"BRUIS 20250420 okt-3"},{"id":221,"rekening":"908","jaar":2025,"maand":10,"datum":"2025-10-28","grootboek_code":"8020","grootboek_naam":"Clubactie Rabobank","bedrag":313.1,"omschrijving":"Rabo ClubSupport"},{"id":222,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-03","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-21.6,"omschrijving":"rente en bankkosten"},{"id":223,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-05","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-8.5,"omschrijving":"Bloemen expositie 2025 Marijke  Donders"},{"id":224,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-05","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-13.49,"omschrijving":"bevestigingstape exp 2025 G Jansen"},{"id":225,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":226,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":227,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":228,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":229,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":230,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":231,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":232,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":233,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":234,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":235,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":236,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":237,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":238,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":239,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":240,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":241,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":242,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-08","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":243,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":244,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":245,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":246,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":247,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":248,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":249,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":250,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":251,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":252,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-09","grootboek_code":"8016","grootboek_naam":"Loterij expositie","bedrag":2.5,"omschrijving":"loterij expo 2025"},{"id":253,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-13","grootboek_code":"1050","grootboek_naam":"Kruisposten","bedrag":27.22,"omschrijving":"foutieve boeking privé Jansen"},{"id":254,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-14","grootboek_code":"1000","grootboek_naam":"Kas","bedrag":34.76,"omschrijving":"storting kasgeld"},{"id":255,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-14","grootboek_code":"1000","grootboek_naam":"Kas","bedrag":2.74,"omschrijving":"storting kasgeld 2,74"},{"id":256,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-14","grootboek_code":"1000","grootboek_naam":"Kas","bedrag":267.26,"omschrijving":"loterij expo 2025 267,26"},{"id":257,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-14","grootboek_code":"1050","grootboek_naam":"Kruisposten","bedrag":-27.22,"omschrijving":"foutieve boeking privé Jansen"},{"id":258,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-17","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-15,"omschrijving":"Vergoeding glasbreuk expo 2025"},{"id":259,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-19","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess vd Spijker"},{"id":260,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-19","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-90,"omschrijving":"20250465 expo 2025 BRUIS"},{"id":261,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-19","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-41.8,"omschrijving":"20250448 nov-1 BRUIS"},{"id":262,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-19","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-16.14,"omschrijving":"diverse aankopen Expo 2025 Lonneke T"},{"id":263,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-19","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-176.2,"omschrijving":"20250478 expo 2025 BRUIS"},{"id":264,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-19","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-5,"omschrijving":"20250464 nov-2 BRUIS"},{"id":265,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-19","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-210,"omschrijving":"vergoeding docent blok 2 Daan vdE"},{"id":266,"rekening":"908","jaar":2025,"maand":11,"datum":"2025-11-20","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-240,"omschrijving":"vergoeding docent blok 2 Angelina v K"},{"id":267,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-02","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-35.23,"omschrijving":"rente en bankkosten"},{"id":268,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-02","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-5,"omschrijving":"20250515 nov-4 BRUIS"},{"id":269,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-11","grootboek_code":"8015","grootboek_naam":"Subsidies en bijdragen","bedrag":1550,"omschrijving":"Incidentele subsitie jeugdatelier Bladel"},{"id":270,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-12","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":8.75,"omschrijving":"Laura van Dijk"},{"id":271,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-16","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-129.28,"omschrijving":"materiaal en begeleiding jeugdat Hennie W"},{"id":272,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-20","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Evi Hoeks"},{"id":273,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-22","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-240,"omschrijving":"Nieuwjaarsbijeenkomst 2026 De Piramide"},{"id":274,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-22","grootboek_code":"4650","grootboek_naam":"Verzekeringen","bedrag":-155.79,"omschrijving":"Quintes aansprakelijkheidsverzekering"},{"id":275,"rekening":"908","jaar":2025,"maand":12,"datum":"2025-12-22","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-37.4,"omschrijving":"20250586 dec-1 BRUIS"},{"id":276,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-01","grootboek_code":"1130","grootboek_naam":"Rabobank .319","bedrag":114.1,"omschrijving":"overboeking rente spaarrekening"},{"id":277,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-02","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Wouters"},{"id":278,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-02","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-5,"omschrijving":"20250496 nov-3 BRUIS"},{"id":279,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-02","grootboek_code":"7030","grootboek_naam":"Verbruiksartikelen atelier","bedrag":-84.89,"omschrijving":"diverse aankopen Lonneke T"},{"id":280,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-02","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-16.05,"omschrijving":"rente en kosten bank"},{"id":281,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Angelina van Kemenade"},{"id":282,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Daan van den Enden"},{"id":283,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Gijsbert Jansen"},{"id":284,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Guus Goossens"},{"id":285,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Lonneke Timmermans"},{"id":286,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Maarten van Dijk"},{"id":287,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Mieke Dirks"},{"id":288,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Nelya Pelymskykh"},{"id":289,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Renée Holtzer"},{"id":290,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Rieky van Iersel"},{"id":291,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-04","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Walter Hendrickx"},{"id":292,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Dorien Reijnders"},{"id":293,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Elly Hovens"},{"id":294,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Elly Verrijt"},{"id":295,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Hanneke de Groot"},{"id":296,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Jeanne van Leeuwen"},{"id":297,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Karin Tonneijk"},{"id":298,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Lein Kloet"},{"id":299,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Maria Verweij-Bekx"},{"id":300,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Marie-Hélène Stokkink"},{"id":301,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Nell Beerends"},{"id":302,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Tonny vd Tillaar"},{"id":303,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Toos Adams"},{"id":304,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-05","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Truus Stadler"},{"id":305,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-06","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Ingrid Spliethof"},{"id":306,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-06","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-253.59,"omschrijving":"0601 mat jeugd DaanvdE"},{"id":307,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-06","grootboek_code":"4300","grootboek_naam":"Huur","bedrag":-835.05,"omschrijving":"Huur en energie oude atelier 11e kw 2025"},{"id":308,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-06","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Mariette van Sambeeck"},{"id":309,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-07","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Harry Jacobs"},{"id":310,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-07","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Marietje de Graaf"},{"id":311,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-07","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Noortje Maas"},{"id":312,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-07","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Zlaty grem"},{"id":313,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-08","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Monique Kuijpers"},{"id":314,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sophie Merks"},{"id":315,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-12","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-716.13,"omschrijving":"20250617 BRUIS afrekening koffie/thee 2025"},{"id":316,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-12","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-7.8,"omschrijving":"20250633 BRUIS"},{"id":317,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Ingrid de Wit"},{"id":318,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Mara van Straaten"},{"id":319,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Waarma"},{"id":320,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Ize Sebregts"},{"id":321,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Evi Gooskens"},{"id":322,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lynn Vernie"},{"id":323,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":50,"omschrijving":"Lola en Coco Pinsen"},{"id":324,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Marit Janssen"},{"id":325,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Emilie van Gestel"},{"id":326,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-13","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Rina Kennis"},{"id":327,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-14","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Annie van Oorschot"},{"id":328,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-14","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Anton Cools"},{"id":329,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-14","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Dick Scheepmaker"},{"id":330,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-14","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Margaret Janssen"},{"id":331,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-14","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vera Tenbult"},{"id":332,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-14","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Jinze van Limpt"},{"id":333,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-14","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Livia Rombouts"},{"id":334,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-15","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Andrea Boullart"},{"id":335,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-15","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Cor van Gool"},{"id":336,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-15","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Hennie Walenberg"},{"id":337,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-15","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Jolanda Spaan"},{"id":338,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-15","grootboek_code":"8015","grootboek_naam":"Subsidies en bijdragen","bedrag":2000,"omschrijving":"Voorschot Subsidie gemeente Bladel"},{"id":339,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-19","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Gerthy Fleskens"},{"id":340,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-19","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":20,"omschrijving":"Bo Peijs"},{"id":341,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-20","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Marijke Donders"},{"id":342,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-20","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Twan Duis"},{"id":343,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sanna van den Borne"},{"id":344,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-21","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Liesbeth Adams"},{"id":345,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Izebeau te Riele"},{"id":346,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-21","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Eva van Deursen"},{"id":347,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-22","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Laura van Dijk"},{"id":348,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-22","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Arianne van Hout"},{"id":349,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-22","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Jeanette Graamans"},{"id":350,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-25","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":100,"omschrijving":"Sofiya Verdonschot"},{"id":351,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-26","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Huub van der Krogt"},{"id":352,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-26","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Peter vd Wouw"},{"id":353,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-28","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-250.69,"omschrijving":"diverse aankopen Lonneke T"},{"id":354,"rekening":"908","jaar":2026,"maand":1,"datum":"2026-01-30","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-20.53,"omschrijving":"mat jeugd Angelina  vK"},{"id":355,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-03","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-32.69,"omschrijving":"rente en kosten bank"},{"id":356,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-04","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":200,"omschrijving":"retour borg de Piramide nieuwjaarsborrel"},{"id":357,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-05","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-57.38,"omschrijving":"mat jeugd Daan vdE"},{"id":358,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-05","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-210,"omschrijving":"begl jeugd Daan vdE"},{"id":359,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-08","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Evi Gooskens"},{"id":360,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-09","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-5.5,"omschrijving":"20250686 BRUIS jan-2"},{"id":361,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-10","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess vd Spijker"},{"id":362,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-13","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-280,"omschrijving":"begl jeugd Angelina vK"},{"id":363,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-19","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-128.75,"omschrijving":"20260719 BRUIS fact jan-3"},{"id":364,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-21","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Karien vd Heijden"},{"id":365,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-21","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":26.45,"omschrijving":"retour statiegeld nieuwjaarsborrel"},{"id":366,"rekening":"908","jaar":2026,"maand":2,"datum":"2026-02-24","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lynn Vernie"},{"id":367,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-02","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vivien Strijbos"},{"id":368,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-02","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-16.82,"omschrijving":"rente en kosten bank"},{"id":369,"rekening":"908","jaar":2026,"maand":3,"datum":"2025-03-06","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Waarma"},{"id":370,"rekening":"908","jaar":2026,"maand":3,"datum":"2025-03-06","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Wouters"},{"id":371,"rekening":"908","jaar":2026,"maand":3,"datum":"2025-03-06","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vera Tenbult"},{"id":372,"rekening":"908","jaar":2026,"maand":3,"datum":"2025-03-06","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sophie Merks"},{"id":373,"rekening":"908","jaar":2026,"maand":3,"datum":"2025-03-06","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Emilie van Gestel"},{"id":374,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-07","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Izebeau te Riele"},{"id":375,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-07","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Marit Janssen"},{"id":376,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-09","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-135,"omschrijving":"20260787 BRUIS mrt-1"},{"id":377,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-09","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Evi Hoeks"},{"id":378,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-09","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Livia Rombouts"},{"id":379,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-09","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Ize Sebregts"},{"id":380,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-10","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":5,"omschrijving":"Sofiya Verdonschot"},{"id":381,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-10","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":105,"omschrijving":"Linda van der Burg"},{"id":382,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-10","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":87.5,"omschrijving":"Bonni Boelhouwers"},{"id":383,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":50,"omschrijving":"Lola en Coco Pinsen"},{"id":384,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":20,"omschrijving":"Noor v Kreij"},{"id":385,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-14","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Bo Peijs"},{"id":386,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-16","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Jinze van Limpt"},{"id":387,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-16","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sanna van den Borne"},{"id":388,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-18","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Eva van Deursen"},{"id":389,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-19","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-22.9,"omschrijving":"20260795 BRUIS mrt-2"},{"id":390,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-20","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-75,"omschrijving":"kadobonnen Jan vd Putten Workshop paletmet"},{"id":391,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-25","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-52.5,"omschrijving":"20260849 BRUIS koffie-thee workshop paletmes"},{"id":392,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-31","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-66.82,"omschrijving":"mat jeugd Daan vdE mrt 26"},{"id":393,"rekening":"908","jaar":2026,"maand":3,"datum":"2026-03-31","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":10,"omschrijving":"Amy van Vessem"},{"id":394,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-20.42,"omschrijving":"rente en kosten bank"},{"id":395,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-03","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-17.85,"omschrijving":"mat jeugd Daan vdE mrt-2 26"},{"id":396,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-07","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Lynn Vernie"},{"id":397,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-07","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Marit Janssen"},{"id":398,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-07","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess vd Spijker"},{"id":399,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-08","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Jinze van Limpt"},{"id":400,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-13","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-9.72,"omschrijving":"mat jeugd Angelina  vK"},{"id":401,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-13","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-280,"omschrijving":"begl. Jeugd Blok 4 Angelina vK"},{"id":402,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-13","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-315,"omschrijving":"begl. Jeugd Blok 4 Daan vdE"},{"id":403,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-20","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-144.5,"omschrijving":"20260911 BRUIS apr-1"},{"id":404,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-20","grootboek_code":"1130","grootboek_naam":"Rabobank .319","bedrag":-2000,"omschrijving":"overboeking naar spaarrekening"},{"id":405,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-20","grootboek_code":"4611","grootboek_naam":"Website","bedrag":-95.59,"omschrijving":"graaggoedonline AVG-pakket"},{"id":406,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-21","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":70,"omschrijving":"Fred Bleijs"},{"id":407,"rekening":"908","jaar":2026,"maand":4,"datum":"2026-04-29","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Ivy Hurkmans"},{"id":408,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-03","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-38.33,"omschrijving":"mat.  Jeugd DaanvdE mei-1"},{"id":409,"rekening":"908","jaar":2026,"maand":5,"datum":"2025-05-04","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-17.13,"omschrijving":"rente en kosten bank"},{"id":410,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-04","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Mick Lepelaars"},{"id":411,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-05","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Sophie Merks"},{"id":412,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-06","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":87.5,"omschrijving":"Noor vd Molengraft"},{"id":413,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Ize Sebregts"},{"id":414,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess Waarma"},{"id":415,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vera Tenbult"},{"id":416,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Emilie van Gestel"},{"id":417,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Evi Gooskens"},{"id":418,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-12","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Vivien Strijbos"},{"id":419,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Nune van der  Heijden"},{"id":420,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-13","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Livia Rombouts"},{"id":421,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-19","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Coco Prinsen"},{"id":422,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-20","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Amy van Vessem"},{"id":423,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-20","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-18.85,"omschrijving":"mat. Jeugd DaanvdE mei-2"},{"id":424,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-20","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-120,"omschrijving":"begl. jeugd Hennie W"},{"id":425,"rekening":"908","jaar":2026,"maand":5,"datum":"2026-05-27","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Eva van Deursen"},{"id":426,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-18.3,"omschrijving":"rente en kosten bank"},{"id":427,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-08","grootboek_code":"7020","grootboek_naam":"Kosten activiteiten","bedrag":-291.65,"omschrijving":"buiten schilderen Voor Anker 6-6-26"},{"id":428,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-08","grootboek_code":"1000","grootboek_naam":"Kas","bedrag":-700,"omschrijving":"Expositieschotten revisie excl bon"},{"id":429,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-08","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-198.5,"omschrijving":"20260973 BRUIS mei-1"},{"id":430,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-08","grootboek_code":"4320","grootboek_naam":"Inventaris atelier","bedrag":-847,"omschrijving":"Expositieschotten revisie excl bon"},{"id":431,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-08","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-19.96,"omschrijving":"mat jeugd Angelina  vK jun-1"},{"id":432,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-08","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Tess vd Spijker"},{"id":433,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-09","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Noor van Kreij"},{"id":434,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-09","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-252,"omschrijving":"begl jeugd Daan vdE"},{"id":435,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-09","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-16.09,"omschrijving":"mat jeugd Daan vdE jun-1"},{"id":436,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-09","grootboek_code":"4005","grootboek_naam":"Schoonmaakkosten atelier","bedrag":-141.58,"omschrijving":"PureHygiene papieren handdoeken"},{"id":437,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-11","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"saanne vd Borne (cash via DaanvdE)"},{"id":438,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-12","grootboek_code":"7035","grootboek_naam":"Verbruiksartikelen jeugdatelier","bedrag":-9.68,"omschrijving":"mat. Jeugd AngelinavK jun-1"},{"id":439,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-12","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":-245,"omschrijving":"begl Jeugd Angelina vK"},{"id":440,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-12","grootboek_code":"7025","grootboek_naam":"Kosten jeugdatelier","bedrag":7,"omschrijving":"begl jeugd Daan vdE € 7 teveel in rekening"},{"id":441,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-14","grootboek_code":"8002","grootboek_naam":"Contributie bijdragen","bedrag":70,"omschrijving":"Anja van Hattem"},{"id":442,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-15","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-108,"omschrijving":"20261068 BRUIS jun-1"},{"id":443,"rekening":"908","jaar":2026,"maand":6,"datum":"2026-06-25","grootboek_code":"4640","grootboek_naam":"Bestuurskosten","bedrag":-29.6,"omschrijving":"20261081 BRUIS fact jun-2"},{"id":444,"rekening":"908","jaar":2026,"maand":7,"datum":"2026-07-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-18.82,"omschrijving":"rente en kosten bank"},{"id":445,"rekening":"908","jaar":2026,"maand":7,"datum":"2026-07-16","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":25,"omschrijving":"Fenna Koolen"},{"id":446,"rekening":"908","jaar":2026,"maand":7,"datum":"2026-07-18","grootboek_code":"8005","grootboek_naam":"Contributie jeugdatelier","bedrag":50,"omschrijving":"Saar en Pien van der Kroon"},{"id":447,"rekening":"908","jaar":2026,"maand":7,"datum":"2026-07-22","grootboek_code":"4205","grootboek_naam":"Kantinekosten","bedrag":-112.25,"omschrijving":"20261129 BRUIS fact jul-1"},{"id":448,"rekening":"319","jaar":2025,"maand":1,"datum":"2025-01-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":139.23,"omschrijving":"rente periode 01-01-24 t/m 31-12-2024"},{"id":449,"rekening":"319","jaar":2025,"maand":1,"datum":"2025-01-01","grootboek_code":"1120","grootboek_naam":"Rabobank .908","bedrag":-139.23,"omschrijving":"overboeking rente spaarrekening"},{"id":450,"rekening":"319","jaar":2025,"maand":10,"datum":"2025-10-06","grootboek_code":"1120","grootboek_naam":"Rabobank .908","bedrag":-1500,"omschrijving":"ophoging saldo"},{"id":451,"rekening":"319","jaar":2026,"maand":1,"datum":"2026-01-01","grootboek_code":"4900","grootboek_naam":"Rente en bankkosten","bedrag":-114.1,"omschrijving":"overboeking rente periode 01-01-25 t/m 31-12-2025"},{"id":452,"rekening":"319","jaar":2026,"maand":1,"datum":"2026-01-01","grootboek_code":"4901","grootboek_naam":"Rente en bankkosten","bedrag":114.1,"omschrijving":"rente periode 01-01-25 t/m 31-12-2026"}];
const SEED_ACCOUNTS = [{"code":"0220","naam":"Algemene reserve"},{"code":"1000","naam":"Kas"},{"code":"1050","naam":"Kruisposten"},{"code":"1120","naam":"Rabobank .908"},{"code":"1130","naam":"Rabobank .319"},{"code":"1500","naam":"Betaalde BTW"},{"code":"1610","naam":"Overlopende passiva"},{"code":"1620","naam":"Overlopende activa"},{"code":"4005","naam":"Schoonmaakkosten atelier"},{"code":"4205","naam":"Kantinekosten"},{"code":"4300","naam":"Huur"},{"code":"4310","naam":"Energiekosten"},{"code":"4318","naam":"Schoonmaakkosten"},{"code":"4320","naam":"Inventaris atelier"},{"code":"4390","naam":"Overige huisvestingskosten"},{"code":"4440","naam":"Algemene ledenvergadering"},{"code":"4611","naam":"Website"},{"code":"4620","naam":"Representatiekosten"},{"code":"4640","naam":"Bestuurskosten"},{"code":"4650","naam":"Verzekeringen"},{"code":"4820","naam":"Notariskosten"},{"code":"4900","naam":"Rente en bankkosten"},{"code":"4901","naam":"Ontvangen rente"},{"code":"7020","naam":"Kosten activiteiten"},{"code":"7025","naam":"Kosten jeugdatelier"},{"code":"7030","naam":"Verbruiksartikelen atelier"},{"code":"7035","naam":"Verbruiksartikelen jeugdatelier"},{"code":"8002","naam":"Contributie bijdragen"},{"code":"8005","naam":"Contributie jeugdatelier"},{"code":"8012","naam":"Opbrengst verkoop papier"},{"code":"8015","naam":"Subsidies en bijdragen"},{"code":"8016","naam":"Loterij expositie"},{"code":"8020","naam":"Clubactie Rabobank"},{"code":"8021","naam":"Spaaracties overig"},{"code":"9000","naam":"Nog uitzoeken"}];
const SEED_BUDGET = {"2026":[{"sectie":"Inkomsten","categorie":"Contributie","bedrag":7760,"notitie":null},{"sectie":"Inkomsten","categorie":"subsidies en bijdragen","bedrag":1500,"notitie":null},{"sectie":"Inkomsten","categorie":"clubactie Rabobank","bedrag":350,"notitie":null},{"sectie":"Inkomsten","categorie":"spaaracties overig","bedrag":200,"notitie":null},{"sectie":"Inkomsten","categorie":"onvoorzien","bedrag":0,"notitie":null},{"sectie":"Uitgaven","categorie":"1500  Betaalde BTW","bedrag":0,"notitie":null},{"sectie":"Uitgaven","categorie":"4005  Schoonmaakkosten atelier","bedrag":200,"notitie":"Papier, zeep, diversen"},{"sectie":"Uitgaven","categorie":"4205  Kantinekosten","bedrag":2175,"notitie":null},{"sectie":"Uitgaven","categorie":"4300  Huur","bedrag":650,"notitie":"kwartaal 1 2025 oude atelier"},{"sectie":"Uitgaven","categorie":"4310  Energiekosten","bedrag":300,"notitie":"kwartaal 1 2025 oude atelier"},{"sectie":"Uitgaven","categorie":"4320  Inventaris atelier","bedrag":300,"notitie":null},{"sectie":"Uitgaven","categorie":"4390  Overige huisvestingskosten","bedrag":100,"notitie":"stelpost + € 25 WOZ oude atelier"},{"sectie":"Uitgaven","categorie":"4440  Algemene ledenvergadering","bedrag":200,"notitie":null},{"sectie":"Uitgaven","categorie":"4611  Website","bedrag":150,"notitie":null},{"sectie":"Uitgaven","categorie":"4620  Representatiekosten","bedrag":250,"notitie":"bloemen, flyers-advertentie"},{"sectie":"Uitgaven","categorie":"4640 Bestuurskosten","bedrag":350,"notitie":null},{"sectie":"Uitgaven","categorie":"4650  Verzekeringen","bedrag":175,"notitie":null},{"sectie":"Uitgaven","categorie":"4820 Notariskosten","bedrag":0,"notitie":null},{"sectie":"Uitgaven","categorie":"4900  Rente en bankkosten","bedrag":350,"notitie":null},{"sectie":"Uitgaven","categorie":"7020 Kosten activiteiten","bedrag":350,"notitie":null},{"sectie":"Uitgaven","categorie":"7025  kosten jeugdatelier","bedrag":2600,"notitie":null},{"sectie":"Uitgaven","categorie":"Onvoorzien","bedrag":500,"notitie":null},{"sectie":"Uitgaven","categorie":"Algemene reserve","bedrag":7842,"notitie":null}],"2027":[{"sectie":"Inkomsten","categorie":"Contributie","bedrag":7470,"notitie":null},{"sectie":"Inkomsten","categorie":"subsidies en bijdragen","bedrag":2000,"notitie":null},{"sectie":"Inkomsten","categorie":"clubactie Rabobank","bedrag":350,"notitie":null},{"sectie":"Inkomsten","categorie":"spaaracties overig","bedrag":200,"notitie":null},{"sectie":"Inkomsten","categorie":"onvoorzien","bedrag":0,"notitie":null},{"sectie":"Uitgaven","categorie":"1500  Betaalde BTW","bedrag":0,"notitie":null},{"sectie":"Uitgaven","categorie":"4005  Schoonmaakkosten atelier","bedrag":220,"notitie":"Papier, zeep, diversen"},{"sectie":"Uitgaven","categorie":"4205  Kantinekosten","bedrag":2280,"notitie":null},{"sectie":"Uitgaven","categorie":"4300  Huur","bedrag":0,"notitie":"kwartaal 1 2025 oude atelier"},{"sectie":"Uitgaven","categorie":"4310  Energiekosten","bedrag":0,"notitie":"kwartaal 1 2025 oude atelier"},{"sectie":"Uitgaven","categorie":"4320  Inventaris atelier","bedrag":2575,"notitie":null},{"sectie":"Uitgaven","categorie":"4390  Overige huisvestingskosten","bedrag":100,"notitie":"stelpost"},{"sectie":"Uitgaven","categorie":"4440  Algemene ledenvergadering","bedrag":250,"notitie":null},{"sectie":"Uitgaven","categorie":"4611  Website","bedrag":500,"notitie":null},{"sectie":"Uitgaven","categorie":"4620  Representatiekosten","bedrag":250,"notitie":"bloemen, flyers-advertentie"},{"sectie":"Uitgaven","categorie":"4640 Bestuurskosten","bedrag":400,"notitie":null},{"sectie":"Uitgaven","categorie":"4650  Verzekeringen","bedrag":200,"notitie":null},{"sectie":"Uitgaven","categorie":"4820 Notariskosten","bedrag":0,"notitie":null},{"sectie":"Uitgaven","categorie":"4900  Rente en bankkosten","bedrag":420,"notitie":null},{"sectie":"Uitgaven","categorie":"7020 Kosten activiteiten","bedrag":900,"notitie":null},{"sectie":"Uitgaven","categorie":"7025  kosten jeugdatelier","bedrag":3900,"notitie":null},{"sectie":"Uitgaven","categorie":"Algemene reserve","bedrag":8305,"notitie":null}]};
const START_BALANCE = {"908":1726.04,"319":7805.0};
/* Berekent start-, mutatie- en eindsaldo voor een specifiek boekjaar (en optioneel rekening
   '908'/'319', of 'alle' voor beide samen). Startsaldo = vast beginbedrag + alle boekingen
   van vóór dat jaar; er wordt nergens een apart startsaldo per jaar opgeslagen, dit wordt
   steeds live afgeleid uit de boekingen. */
function saldoBoekjaar(tx, rekening, jaar) {
  const rekeningen = rekening === 'alle' ? Object.keys(START_BALANCE) : [rekening];
  let startsaldo = 0, mutaties = 0;
  rekeningen.forEach(r => {
    startsaldo += (START_BALANCE[r] || 0) + tx.filter(t => t.rekening === r && t.jaar < jaar).reduce((s, t) => s + Number(t.bedrag || 0), 0);
    mutaties += tx.filter(t => t.rekening === r && t.jaar === jaar).reduce((s, t) => s + Number(t.bedrag || 0), 0);
  });
  return { startsaldo, mutaties, eindsaldo: startsaldo + mutaties };
}
const SEED_YEARS = [2025,2026,2027];

const APP_VERSIE = '01-09-2026';
const SEED_SLOTS = ["ma 11.00 - 16.00","di 10.00 - 16.00","wo 09.00 - 12.30","wo 19.00 - 22.00","do 09.30 - 16.00","do 19.00 - 22.00"];
const SEED_AGENDAPUNTEN_VOORAF = ["Opening", "Mededelingen", "Vaststellen agenda", "Notulen vorige vergadering"];
const SEED_AGENDAPUNTEN_AFSLUITEND = ["Rondvraag", "Sluiting"];

/* Welke tabbladen mag iemand BEWERKEN, op basis van hun functie (Dashboard en Rapportage
   zijn altijd overal zichtbaar/exporteerbaar — dat zijn geen bewerkacties).
   Herken je hier je eigen bestuursfuncties niet in, vraag dan gewoon om deze lijst aan te passen. */
const EDITEERBARE_TABS = [
  { id: 'leden', label: 'Leden' }, { id: 'workshops', label: 'Workshops' },
  { id: 'vergaderingen', label: 'Vergaderingen' }, { id: 'financien', label: 'Financiën' },
  { id: 'begroting', label: 'Begroting' },
];
const SEED_ROLPERMISSIES = [
  { id: 1, patroon: 'voorzitter', tabs: ['leden', 'workshops', 'vergaderingen', 'financien', 'begroting'] },
  { id: 2, patroon: 'penningmeester', tabs: ['financien', 'begroting'] },
  { id: 3, patroon: 'secretaris', tabs: ['vergaderingen', 'leden'] },
  { id: 4, patroon: 'ledenadministratie', tabs: ['leden', 'workshops'] },
];
function bewerkbareTabs(functie, regels) {
  if (!functie) return [];
  const regel = regels.find(r => functie.toLowerCase().includes(r.patroon.toLowerCase()));
  return regel ? regel.tabs : [];
}
const MONTH_NAMES = ["jan","feb","mrt","apr","mei","jun","jul","aug","sep","okt","nov","dec"];

/* ---------------------------- design tokens ---------------------------- */
const C = {
  paper: '#F5F8FC',
  paperDim: '#E9F1F9',
  ink: '#13244A',
  inkSoft: '#5A6B8C',
  clay: '#0091D6',
  clayDeep: '#0B3F8F',
  sage: '#0E7A5B',
  sageDeep: '#0A5E46',
  ochre: '#D9A300',
  rose: '#9C3568',
  border: '#DCE6F2',
  card: '#FFFFFF',
};
const LOGO_URI = 'data:image/png;base64,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';

/* ------------------------------ helpers -------------------------------- */
function euro(n) {
  const v = Number(n) || 0;
  return v.toLocaleString('nl-NL', { style: 'currency', currency: 'EUR' });
}
function fmtDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  return d.toLocaleDateString('nl-NL', { day: '2-digit', month: '2-digit', year: 'numeric' });
}
/* ==== Twee-factor authenticatie (TOTP, RFC 6238) — puur client-side via Web Crypto ====
   Geen externe bibliotheek nodig: base32-encodering + HMAC-SHA1 via crypto.subtle. */
const BASE32_ALFABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
function genereerTotpSecret() {
  const bytes = crypto.getRandomValues(new Uint8Array(20));
  let bits = '';
  for (const b of bytes) bits += b.toString(2).padStart(8, '0');
  let secret = '';
  for (let i = 0; i + 5 <= bits.length; i += 5) secret += BASE32_ALFABET[parseInt(bits.slice(i, i + 5), 2)];
  return secret;
}
function base32Decode(str) {
  const clean = str.toUpperCase().replace(/[^A-Z2-7]/g, '');
  let bits = '';
  for (const ch of clean) bits += BASE32_ALFABET.indexOf(ch).toString(2).padStart(5, '0');
  const bytes = [];
  for (let i = 0; i + 8 <= bits.length; i += 8) bytes.push(parseInt(bits.slice(i, i + 8), 2));
  return new Uint8Array(bytes);
}
async function hotp(secretBytes, counter) {
  const counterBytes = new Uint8Array(8);
  let c = BigInt(counter);
  for (let i = 7; i >= 0; i--) { counterBytes[i] = Number(c & 0xffn); c >>= 8n; }
  const key = await crypto.subtle.importKey('raw', secretBytes, { name: 'HMAC', hash: 'SHA-1' }, false, ['sign']);
  const sig = new Uint8Array(await crypto.subtle.sign('HMAC', key, counterBytes));
  const offset = sig[sig.length - 1] & 0xf;
  const code = ((sig[offset] & 0x7f) << 24) | ((sig[offset + 1] & 0xff) << 16) | ((sig[offset + 2] & 0xff) << 8) | (sig[offset + 3] & 0xff);
  return String(code % 1000000).padStart(6, '0');
}
async function totpHuidigeCode(secret) {
  return hotp(base32Decode(secret), Math.floor(Date.now() / 1000 / 30));
}
async function verifieerTotp(secret, code) {
  const schoon = (code || '').replace(/\D/g, '');
  if (schoon.length !== 6) return false;
  const stap = Math.floor(Date.now() / 1000 / 30);
  const bytes = base32Decode(secret);
  // Sta ±1 tijdstap (30s) toe voor klokverschillen tussen apparaat en telefoon.
  for (const offset of [0, -1, 1]) {
    if (await hotp(bytes, stap + offset) === schoon) return true;
  }
  return false;
}
/* Wachtwoordeisen — voorbereid voor later gebruik, nog niet actief in de huidige inlogflow
   tenzij ingeschakeld via Instellingen → Beveiliging. */
function wachtwoordVoldoetAanEisen(w) {
  return !!w && w.length >= 8 && /[A-Z]/.test(w) && /[a-z]/.test(w) && /[0-9]/.test(w) && /[^A-Za-z0-9]/.test(w);
}

/* ==== Gehashte opslag van pincodes/wachtwoorden (PBKDF2-SHA256, uniek zout per persoon) ====
   Er wordt nooit meer het echte wachtwoord opgeslagen, alleen een onomkeerbare hash + het
   zout dat nodig is om diezelfde hash later opnieuw te kunnen berekenen ter controle.
   Bevat ook een soepele overgang: een nog niet-gehashte (oude) pincode wordt bij de eerst-
   volgende succesvolle login automatisch alsnog gehasht opgeslagen. */
function bytesToHex(bytes) { return Array.from(bytes).map(b => b.toString(16).padStart(2, '0')).join(''); }
function hexToBytes(hex) { const arr = new Uint8Array(hex.length / 2); for (let i = 0; i < arr.length; i++) arr[i] = parseInt(hex.substr(i * 2, 2), 16); return arr; }
async function hashWachtwoord(wachtwoord, saltHex) {
  const salt = saltHex ? hexToBytes(saltHex) : crypto.getRandomValues(new Uint8Array(16));
  const saltHexOut = saltHex || bytesToHex(salt);
  const keyMateriaal = await crypto.subtle.importKey('raw', new TextEncoder().encode(wachtwoord), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt, iterations: 100000, hash: 'SHA-256' }, keyMateriaal, 256);
  return { hash: bytesToHex(new Uint8Array(bits)), salt: saltHexOut };
}
async function verifieerWachtwoord(wachtwoord, opgeslagen) {
  if (typeof opgeslagen === 'string') return opgeslagen === wachtwoord; // nog niet-gehashte (oude) pincode
  if (!opgeslagen || !opgeslagen.hash || !opgeslagen.salt) return false;
  const { hash } = await hashWachtwoord(wachtwoord, opgeslagen.salt);
  return hash === opgeslagen.hash;
}

/* ==== QR-code generator (ISO/IEC 18004) - byte mode, foutcorrectieniveau L, een RS-blok,
   versies 1 t/m 5. Puur JavaScript, geen externe bibliotheek. Zelf geverifieerd door een
   onafhankelijk geschreven decoder die de matrix terugleest en exact de brontekst teruggeeft. */
const QR_GF_EXP = new Array(512);
const QR_GF_LOG = new Array(256);
(function initQrGF() {
  let x = 1;
  for (let i = 0; i < 255; i++) {
    QR_GF_EXP[i] = x; QR_GF_LOG[x] = i;
    x <<= 1;
    if (x & 0x100) x ^= 0x11d;
  }
  for (let i = 255; i < 512; i++) QR_GF_EXP[i] = QR_GF_EXP[i - 255];
})();
function qrGfMul(a, b) { return (a === 0 || b === 0) ? 0 : QR_GF_EXP[QR_GF_LOG[a] + QR_GF_LOG[b]]; }
function qrRsGenerator(degree) {
  let poly = [1];
  for (let i = 0; i < degree; i++) {
    const next = new Array(poly.length + 1).fill(0);
    for (let j = 0; j < poly.length; j++) {
      next[j] ^= qrGfMul(poly[j], 1);
      next[j + 1] ^= qrGfMul(poly[j], QR_GF_EXP[i]);
    }
    poly = next;
  }
  return poly;
}
function qrRsEncode(dataCw, eccCount) {
  const gen = qrRsGenerator(eccCount);
  const res = dataCw.slice();
  for (let i = 0; i < eccCount; i++) res.push(0);
  for (let i = 0; i < dataCw.length; i++) {
    const coef = res[i];
    if (coef === 0) continue;
    for (let j = 0; j < gen.length; j++) res[i + j] ^= qrGfMul(gen[j], coef);
  }
  return res.slice(dataCw.length);
}
const QR_VERSIONS = {
  1: { size: 21, total: 26, ecc: 7, data: 19 },
  2: { size: 25, total: 44, ecc: 10, data: 34 },
  3: { size: 29, total: 70, ecc: 15, data: 55 },
  4: { size: 33, total: 100, ecc: 20, data: 80 },
  5: { size: 37, total: 134, ecc: 26, data: 108 },
};
function qrEncodeByteMode(str, version) {
  const v = QR_VERSIONS[version];
  const bytes = Array.from(new TextEncoder().encode(str));
  let bits = '0100' + bytes.length.toString(2).padStart(8, '0');
  for (const b of bytes) bits += b.toString(2).padStart(8, '0');
  const totalBits = v.data * 8;
  if (bits.length > totalBits) return null;
  bits += '0'.repeat(Math.min(4, totalBits - bits.length));
  while (bits.length % 8 !== 0) bits += '0';
  const codewords = [];
  for (let i = 0; i < bits.length; i += 8) codewords.push(parseInt(bits.slice(i, i + 8), 2));
  const padBytes = [0xec, 0x11];
  let p = 0;
  while (codewords.length < v.data) { codewords.push(padBytes[p % 2]); p++; }
  return codewords;
}
function qrFormatBits() {
  let data = 0b01000;
  let bch = data << 10;
  const gen = 0b10100110111;
  for (let i = 4; i >= 0; i--) if (bch & (1 << (i + 10))) bch ^= gen << i;
  let full = ((data << 10) | bch) ^ 0b101010000010010;
  return full.toString(2).padStart(15, '0');
}
function qrBuildMatrix(version, dataAndEcc) {
  const size = QR_VERSIONS[version].size;
  const mat = Array.from({ length: size }, () => new Array(size).fill(null));
  const reserved = Array.from({ length: size }, () => new Array(size).fill(false));
  function set(r, c, val) { if (r < 0 || c < 0 || r >= size || c >= size) return; mat[r][c] = val; reserved[r][c] = true; }
  function placeFinder(r, c) {
    for (let i = -1; i <= 7; i++) for (let j = -1; j <= 7; j++) {
      const rr = r + i, cc = c + j;
      if (rr < 0 || cc < 0 || rr >= size || cc >= size) continue;
      let val;
      if (i === -1 || i === 7 || j === -1 || j === 7) val = 0;
      else if (i === 0 || i === 6 || j === 0 || j === 6) val = 1;
      else if (i >= 2 && i <= 4 && j >= 2 && j <= 4) val = 1;
      else val = 0;
      set(rr, cc, val);
    }
  }
  placeFinder(0, 0); placeFinder(0, size - 7); placeFinder(size - 7, 0);
  for (let i = 8; i < size - 8; i++) { set(6, i, i % 2 === 0 ? 1 : 0); set(i, 6, i % 2 === 0 ? 1 : 0); }
  set(4 * version + 9, 8, 1);
  const alignCenters = { 2: [6, 18], 3: [6, 22], 4: [6, 26], 5: [6, 30] };
  if (alignCenters[version]) {
    const centers = alignCenters[version];
    for (const r of centers) for (const c of centers) {
      if ((r === 6 && c === 6) || (r === 6 && c === size - 7) || (r === size - 7 && c === 6)) continue;
      for (let i = -2; i <= 2; i++) for (let j = -2; j <= 2; j++) set(r + i, c + j, (Math.max(Math.abs(i), Math.abs(j)) !== 1) ? 1 : 0);
    }
  }
  for (let i = 0; i < 9; i++) { reserved[8][i] = true; reserved[i][8] = true; }
  for (let i = 0; i < 8; i++) { reserved[8][size - 1 - i] = true; reserved[size - 1 - i][8] = true; }

  let bits = '';
  for (const cw of dataAndEcc) bits += cw.toString(2).padStart(8, '0');
  let bitIdx = 0, dir = -1, col = size - 1;
  while (col > 0) {
    if (col === 6) col--;
    for (let i = 0; i < size; i++) {
      const row = dir === -1 ? size - 1 - i : i;
      for (const c of [col, col - 1]) {
        if (reserved[row][c]) continue;
        const bit = bitIdx < bits.length ? parseInt(bits[bitIdx]) : 0;
        bitIdx++;
        mat[row][c] = ((row + c) % 2 === 0) ? (bit ^ 1) : bit;
      }
    }
    dir *= -1; col -= 2;
  }
  const fmt = qrFormatBits().split('').map(Number);
  const topLeft = [[8, 0], [8, 1], [8, 2], [8, 3], [8, 4], [8, 5], [8, 7], [8, 8], [7, 8], [5, 8], [4, 8], [3, 8], [2, 8], [1, 8], [0, 8]];
  for (let i = 0; i < 15; i++) { const [r, c] = topLeft[i]; mat[r][c] = fmt[i]; }
  for (let i = 0; i < 8; i++) mat[8][size - 1 - i] = fmt[i];
  for (let i = 0; i < 7; i++) mat[size - 1 - i][8] = fmt[8 + i];
  return mat;
}
function genereerQrMatrix(tekst) {
  for (const v of [1, 2, 3, 4, 5]) {
    const dataCw = qrEncodeByteMode(tekst, v);
    if (dataCw) return qrBuildMatrix(v, dataCw.concat(qrRsEncode(dataCw, QR_VERSIONS[v].ecc)));
  }
  return null;
}
function QrCode({ tekst, teksten, size = 176 }) {
  const kandidaten = teksten || [tekst];
  const mat = useMemo(() => {
    for (const t of kandidaten) {
      const m = genereerQrMatrix(t);
      if (m) return m;
    }
    return null;
  }, [JSON.stringify(kandidaten)]);
  if (!mat) return null;
  const n = mat.length;
  const cell = size / (n + 2);
  let path = '';
  for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) {
    if (mat[r][c] === 1) path += `M${(c + 1) * cell},${(r + 1) * cell}h${cell}v${cell}h${-cell}z`;
  }
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ background: '#fff', borderRadius: 8 }}>
      <path d={path} fill="#000" />
    </svg>
  );
}

function fullName(p) {
  if (p.naam) return p.naam;
  return [p.voornaam, p.tussenvoegsel, p.achternaam].filter(Boolean).join(' ');
}
function uid(list) {
  return (list.reduce((m, x) => Math.max(m, x.id || 0), 0)) + 1;
}
function telLink(tel) {
  return `tel:${(tel || '').replace(/[^0-9+]/g, '')}`;
}
function waNummer(tel) {
  let n = (tel || '').replace(/[^0-9+]/g, '');
  if (n.startsWith('+')) n = n.slice(1);
  else if (n.startsWith('0')) n = '31' + n.slice(1);
  return n;
}
function waLink(tel, tekst) {
  const n = waNummer(tel);
  return `https://wa.me/${n}${tekst ? '?text=' + encodeURIComponent(tekst) : ''}`;
}
function kopieerFallback(tekst) {
  try {
    const el = document.createElement('textarea');
    el.value = tekst;
    el.style.position = 'fixed';
    el.style.left = '-9999px';
    document.body.appendChild(el);
    el.focus();
    el.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(el);
    return ok ? Promise.resolve() : Promise.reject(new Error('execCommand mislukt'));
  } catch (e) {
    return Promise.reject(e);
  }
}
function kopieerNaarKlembord(tekst) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(tekst).catch(() => kopieerFallback(tekst));
  }
  return kopieerFallback(tekst);
}
/* tel:-links worden door de beveiligde omgeving van dit artifact geblokkeerd wanneer ze
   rechtstreeks vanuit de app worden aangeroepen — net als eerder bij printen. Een apart,
   niet-afgeschermd tabblad dat zelf naar tel: navigeert werkt wél (zelfde truc als de
   Word/PDF-export). */
function belViaNieuwTabblad(tel) {
  const link = telLink(tel);
  const html = `<!DOCTYPE html><html lang="nl"><head><meta charset="utf-8" /><title>Bellen…</title></head>
<body style="font-family:Arial,sans-serif;padding:32px;text-align:center;color:#13244A;">
<p style="font-size:15px;">Dit tabblad probeert nu het bellen te starten…</p>
<p><a href="${link}" style="font-size:18px;color:#0091D6;">Gebeurt er niets? Tik/klik hier om te bellen</a></p>
<script>window.location.href = ${JSON.stringify(link)};</script>
</body></html>`;
  const blob = new Blob([html], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  const win = window.open(url, '_blank');
  if (!win) window.location.href = link;
  setTimeout(() => URL.revokeObjectURL(url), 8000);
}
function CopyButton({ text, title }) {
  const [gekopieerd, setGekopieerd] = useState(false);
  const [mislukt, setMislukt] = useState(false);
  function kopieer(e) {
    e.preventDefault();
    e.stopPropagation();
    kopieerNaarKlembord(text).then(() => {
      setGekopieerd(true);
      setMislukt(false);
      setTimeout(() => setGekopieerd(false), 1500);
    }).catch(() => {
      setMislukt(true);
      setTimeout(() => setMislukt(false), 2500);
    });
  }
  return (
    <button type="button" onClick={kopieer} title={mislukt ? 'Kopiëren niet gelukt — selecteer het nummer handmatig' : (title || 'Kopieer nummer')} className="p-1 rounded active:bg-black/10" style={{ color: gekopieerd ? '#0A5E46' : (mislukt ? '#9C3568' : '#5A6B8C'), touchAction: 'manipulation' }}>
      {gekopieerd ? <CheckCircle2 size={14} /> : <Copy size={14} />}
    </button>
  );
}
function ContactActies({ telefoon, naam, toonNummer }) {
  if (!telefoon) return null;
  const btnStyle = { touchAction: 'manipulation' };
  return (
    <span className="inline-flex items-center gap-1" onClick={e => e.stopPropagation()}>
      {toonNummer && (
        <span title="Tik en houd vast (of dubbelklik) om handmatig te selecteren en te kopiëren"
          style={{ userSelect: 'text', WebkitUserSelect: 'text', color: '#5A6B8C' }} className="text-xs">
          {telefoon}
        </span>
      )}
      <button type="button" onClick={() => belViaNieuwTabblad(telefoon)} title={`Bel ${naam}`} className="p-1 rounded active:bg-black/10" style={{ color: '#0091D6', ...btnStyle }}><Phone size={14} /></button>
      <a href={waLink(telefoon)} target="_blank" rel="noreferrer" title={`WhatsApp ${naam}`} className="p-1 rounded active:bg-black/10" style={{ color: '#0A5E46', ...btnStyle }}><MessageCircle size={14} /></a>
      <CopyButton text={telefoon} title={`Kopieer nummer van ${naam}`} />
    </span>
  );
}
function gbLabel(t) {
  return [t.grootboek_code, t.grootboek_naam].filter(Boolean).join('  ');
}
/* Sommige begrotingscategorieën (vooral aan de inkomstenkant, zoals "Contributie" of
   "subsidies en bijdragen") hebben geen grootboekcode vóór de naam staan. Voor die
   gevallen matchen we op trefwoorden naar de bijbehorende grootboekrekening(en), zodat
   het werkelijke bedrag toch meetelt. Retourneert een array van codes (kan leeg zijn). */
const SEED_BEGROTING_KOPPELINGEN = [
  { id: 1, trefwoord: 'aanpassing', codes: '', uitsluiten: true },
  { id: 2, trefwoord: 'resulta', codes: '', uitsluiten: true },
  { id: 3, trefwoord: 'reserve', codes: '', uitsluiten: true },
  { id: 4, trefwoord: 'contributie', codes: '8002,8005', uitsluiten: false },
  { id: 5, trefwoord: 'subsidie', codes: '8015', uitsluiten: false },
  { id: 6, trefwoord: 'clubactie', codes: '8020', uitsluiten: false },
  { id: 7, trefwoord: 'rabobank', codes: '8020', uitsluiten: false },
  { id: 8, trefwoord: 'spaaractie', codes: '8021', uitsluiten: false },
];
/* Koppelt een begrotingscategorie aan grootboekrekening(en) om het "werkelijk"-bedrag te
   kunnen berekenen. Begint de categorie al met een code (bv. "4300 Huur"), dan wordt die
   direct gebruikt. Anders wordt de bewerkbare koppeltabel (Instellingen \u2192 Begroting-koppelingen)
   doorzocht op het eerste passende trefwoord. */
function budgetGbCodes(categorie, koppelingen) {
  const tekst = (categorie || '').trim();
  const m = /^(\d{3,4})/.exec(tekst);
  if (m) return [m[1]];
  const laag = tekst.toLowerCase();
  for (const k of (koppelingen || SEED_BEGROTING_KOPPELINGEN)) {
    if (k.trefwoord && laag.includes(k.trefwoord.toLowerCase())) {
      if (k.uitsluiten) return [];
      return (k.codes || '').split(',').map(c => c.trim()).filter(Boolean);
    }
  }
  return [];
}

/* ------------------------- persistence hook ----------------------------- */
/* Eén gedeelde wachtrij voor ALLE opslagacties in de app (over alle useStored-sleutels heen).
   Zonder dit vuurden gelijktijdige wijzigingen (bv. een boeking + een logboekregel) hun opslag
   tegelijk af, wat de opslag-API overbelastte en tot valse "opslaan mislukt"-meldingen leidde. */
let opslagWachtrij = Promise.resolve();
function planOpslag(taak) {
  opslagWachtrij = opslagWachtrij.then(async () => {
    try { await taak(); } catch (e) { /* taak rapporteert zijn eigen resultaat */ }
    await new Promise(r => setTimeout(r, 180));
  });
  return opslagWachtrij;
}

/* =========================================================================
   OPSLAG-ADAPTER — DE ENIGE PLEK DIE MET DE RUWE OPSLAG-API PRAAT
   ---------------------------------------------------------------------
   Bij een toekomstige migratie naar een externe database (buiten Claude) hoeven
   alléén deze twee functies (opslagLezen / opslagSchrijven) aangepast te worden
   om window.storage.get/set te vervangen door bijvoorbeeld fetch()-aanroepen naar
   een externe API. Alle andere code in de app — de useStored-hook, de sessie-
   "hartslag" (wie is er nog actief) en de back-up-restore-check — loopt via deze
   twee functies en hoeft dus NIET aangepast te worden.

   Alle tabellen (opslagsleutels) die de app gebruikt, staan hieronder in STORAGE_KEYS
   als naslag — dat is het volledige "schema" dat een externe database zou moeten krijgen.
========================================================================= */
const API_SECRET = import.meta.env.VITE_API_SECRET || '';
function apiHeaders(extra) {
  const h = { ...extra };
  if (API_SECRET) h['X-Api-Key'] = API_SECRET;
  return h;
}
async function opslagLezen(key) {
  const res = await fetch(`/api/opslag/${encodeURIComponent(key)}`, { headers: apiHeaders() });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Opslag lezen mislukt (${res.status})`);
  const data = await res.json();
  return { key, value: data.waarde };
}
async function opslagSchrijven(key, waarde) {
  const res = await fetch(`/api/opslag/${encodeURIComponent(key)}`, {
    method: 'PUT',
    headers: apiHeaders({ 'Content-Type': 'application/json' }),
    body: JSON.stringify({ waarde }),
  });
  if (!res.ok) throw new Error(`Opslag schrijven mislukt (${res.status})`);
  return { key, value: waarde };
}
const STORAGE_KEYS = {
  leden: 'bladels:leden', workshops: 'bladels:workshops', workshopinschrijvingen: 'bladels:workshopinschrijvingen',
  transacties: 'bladels:transacties', rekeningen: 'bladels:rekeningen', begrotingkoppelingen: 'bladels:begrotingkoppelingen',
  dagdelen: 'bladels:dagdelen', workshopsoorten: 'bladels:workshopsoorten',
  agendapuntenvooraf: 'bladels:agendapuntenvooraf', agendapuntenafsluitend: 'bladels:agendapuntenafsluitend',
  begroting: 'bladels:begroting', boekjaren: 'bladels:boekjaren', vergaderingen: 'bladels:vergaderingen',
  actielijst: 'bladels:actielijst', contributies: 'bladels:contributies', pins: 'bladels:pins',
  beveiliging: 'bladels:beveiliging', standaarden: 'bladels:standaarden', workshopsortering: 'bladels:workshopsortering',
  tfaSecrets: 'bladels:tfa-secrets', tfaVertrouwd: 'bladels:tfa-vertrouwd', rolpermissies: 'bladels:rolpermissies',
  logboek: 'bladels:logboek', prullenbak: 'bladels:prullenbak', sessies: 'bladels:sessies',
};

function useStored(key, seed, shared = false, onSaved) {
  const [data, setData] = useState(seed);
  const [ready, setReady] = useState(false);
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await opslagLezen(key, shared);
        if (!cancelled && res && res.value) setData(JSON.parse(res.value));
      } catch (e) {
        /* geen opgeslagen versie -> seed blijft actief */
      } finally {
        if (!cancelled) setReady(true);
      }
    })();
    return () => { cancelled = true; };
  }, [key]);

  const persist = useCallback((next) => {
    setData(next);
    planOpslag(async () => {
      const json = JSON.stringify(next);
      try {
        await opslagSchrijven(key, json, shared);
        if (onSaved) onSaved(true);
      } catch (e) {
        // Eerste poging mislukt (vaak een tijdelijke hik) -> na een korte pauze nog één keer proberen
        // voordat we de gebruiker lastigvallen met een foutmelding.
        await new Promise(r => setTimeout(r, 600));
        try {
          await opslagSchrijven(key, json, shared);
          if (onSaved) onSaved(true);
        } catch (e2) {
          console.error('opslaan mislukt voor', key, e2);
          if (onSaved) onSaved(false);
        }
      }
    });
  }, [key]);

  return [data, persist, ready];
}

/* ------------------------------ UI atoms -------------------------------- */
function Card({ children, style, className = '', onClick }) {
  return (
    <div className={`rounded-xl border shadow-sm ${className}`} onClick={onClick}
      style={{ background: C.card, borderColor: C.border, ...style }}>
      {children}
    </div>
  );
}

function Badge({ children, tone = 'sage' }) {
  const map = {
    sage: { bg: '#E9F0E6', fg: C.sageDeep },
    clay: { bg: '#F5E3D6', fg: C.clayDeep },
    ochre: { bg: '#F6EBD3', fg: '#7A5A17' },
    rose: { bg: '#F3E0E0', fg: C.rose },
    muted: { bg: '#EFEAE0', fg: C.inkSoft },
  };
  const t = map[tone] || map.sage;
  return (
    <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium"
      style={{ background: t.bg, color: t.fg }}>
      {children}
    </span>
  );
}

function Btn({ children, onClick, tone = 'clay', type = 'button', size = 'md', icon: Icon, disabled, title }) {
  const tones = {
    clay: { bg: C.clay, fg: '#fff', hover: C.clayDeep },
    sage: { bg: C.sage, fg: '#fff', hover: C.sageDeep },
    ghost: { bg: 'transparent', fg: C.ink, hover: C.paperDim },
    outline: { bg: 'transparent', fg: C.clay, hover: '#F5E3D6' },
    danger: { bg: 'transparent', fg: C.rose, hover: '#F3E0E0' },
    dangerSolid: { bg: C.rose, fg: '#fff', hover: '#7E2A54' },
  };
  const t = tones[tone];
  const pad = size === 'sm' ? 'px-2.5 py-1 text-xs' : 'px-3.5 py-2 text-sm';
  return (
    <button type={type} disabled={disabled} onClick={onClick} title={title}
      className={`inline-flex items-center gap-1.5 rounded-lg font-medium transition-colors border ${pad} ${disabled ? 'opacity-50 cursor-not-allowed' : ''}`}
      style={{ background: t.bg, color: t.fg, borderColor: tone === 'outline' ? C.clay : (tone === 'ghost' ? C.border : 'transparent') }}
      onMouseEnter={e => { if (!disabled) e.currentTarget.style.background = t.hover; }}
      onMouseLeave={e => { if (!disabled) e.currentTarget.style.background = t.bg; }}
    >
      {Icon && <Icon size={14} />}
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="block text-xs font-medium mb-1" style={{ color: C.inkSoft }}>{label}</span>
      {children}
    </label>
  );
}
const inputCls = "w-full rounded-lg border px-3 py-1.5 text-sm outline-none focus:ring-2";
const inputStyle = { borderColor: C.border, background: '#fff' };

/* Voor vrije tekstvelden die anders bij ELKE toetsaanslag zouden opslaan (bv. notulen, bedragen).
   Toont wijzigingen meteen lokaal, maar stuurt pas na een korte pauze (of bij focus verlaten) de
   opslag door — dat voorkomt overbelasting van de opslag-API tijdens het typen. */
function DebouncedField({ value, onCommit, delay = 700, textarea, disabled, className, style, placeholder, type }) {
  const [local, setLocal] = useState(value || '');
  const timer = useRef(null);
  const elRef = useRef(null);
  useEffect(() => { setLocal(value || ''); }, [value]);
  function groei() {
    if (!textarea || !elRef.current) return;
    elRef.current.style.height = 'auto';
    elRef.current.style.height = `${elRef.current.scrollHeight}px`;
  }
  useEffect(() => { groei(); }, [local]);
  function handleChange(e) {
    const v = e.target.value;
    setLocal(v);
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => onCommit(v), delay);
  }
  function commitNow() {
    if (timer.current) { clearTimeout(timer.current); timer.current = null; }
    onCommit(local);
  }
  const Tag = textarea ? 'textarea' : 'input';
  return (
    <Tag
      ref={elRef}
      value={local}
      disabled={disabled}
      placeholder={placeholder}
      type={textarea ? undefined : (type || 'text')}
      rows={textarea ? 3 : undefined}
      onChange={handleChange}
      onBlur={commitNow}
      className={className || inputCls}
      style={style || { ...inputStyle, ...(textarea ? { overflow: 'hidden', resize: 'none', minHeight: '4.5em' } : {}) }}
    />
  );
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start sm:items-center justify-center p-3 sm:p-6 overflow-y-auto"
      style={{ background: 'rgba(42,33,24,0.45)' }} onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className={`w-full ${wide ? 'max-w-2xl' : 'max-w-md'} rounded-2xl shadow-xl my-6`}
        style={{ background: C.card, border: `1px solid ${C.border}` }}>
        <div className="flex items-center justify-between px-5 py-4 border-b" style={{ borderColor: C.border }}>
          <h3 className="font-semibold" style={{ fontFamily: 'Fraunces, serif', color: C.ink, fontSize: 18 }}>{title}</h3>
          <button onClick={onClose} className="rounded-full p-1" style={{ color: C.inkSoft }}><X size={18} /></button>
        </div>
        <div className="px-5 py-4">{children}</div>
      </div>
    </div>
  );
}

function ConfirmModal({ title = 'Weet je het zeker?', message, confirmLabel = 'Verwijderen', onConfirm, onCancel }) {
  return (
    <Modal title={title} onClose={onCancel}>
      <p className="text-sm" style={{ color: C.ink }}>{message}</p>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onCancel}>Annuleren</Btn>
        <Btn tone="dangerSolid" onClick={onConfirm}>{confirmLabel}</Btn>
      </div>
    </Modal>
  );
}

function LoginScreen({ members, pins, setPins, beveiliging, tfaSecrets, setTfaSecrets, tfaVertrouwd, setTfaVertrouwd, onLogin }) {
  const kandidaten = members.filter(m => m.functie && m.functie.trim()).sort((a, b) => fullName(a).localeCompare(fullName(b)));
  const [naam, setNaam] = useState('');
  const [pin, setPin] = useState('');
  const [pin2, setPin2] = useState('');
  const [fout, setFout] = useState('');
  const [resetMode, setResetMode] = useState(false);
  const [stap, setStap] = useState('pincode'); // 'pincode' | 'tfa-setup' | 'tfa-verify'
  const [nieuwSecret, setNieuwSecret] = useState('');
  const [tfaCode, setTfaCode] = useState('');
  const [bezig, setBezig] = useState(false);

  const gekozen = kandidaten.find(m => fullName(m) === naam);
  const heeftAlPin = naam && pins[naam];
  const isNieuw = naam && (!heeftAlPin || resetMode);
  const wachtwoordLabel = beveiliging.wachtwoordEisen ? 'wachtwoord' : 'pincode';

  function tfaVertrouwdNog(persoonNaam) {
    const tot = tfaVertrouwd[persoonNaam];
    return !!tot && new Date(tot).getTime() > Date.now();
  }
  function markeerVertrouwd(persoonNaam) {
    const tot = new Date(Date.now() + beveiliging.tfaVertrouwensdagen * 24 * 60 * 60 * 1000).toISOString();
    setTfaVertrouwd({ ...tfaVertrouwd, [persoonNaam]: tot });
  }

  function na2faControle() {
    if (!beveiliging.tfaVerplicht || tfaVertrouwdNog(naam)) { onLogin(gekozen); return; }
    if (tfaSecrets[naam] && tfaSecrets[naam].bevestigd) {
      setStap('tfa-verify');
    } else {
      const secret = genereerTotpSecret();
      setNieuwSecret(secret);
      setStap('tfa-setup');
    }
  }

  async function submitPincode() {
    setFout('');
    if (!naam) { setFout('Kies eerst je naam.'); return; }
    if (isNieuw) {
      if (beveiliging.wachtwoordEisen) {
        if (!wachtwoordVoldoetAanEisen(pin)) { setFout('Wachtwoord moet minimaal 8 tekens bevatten, met een hoofdletter, kleine letter, cijfer en bijzonder teken.'); return; }
      } else if (pin.length < 4) { setFout('Kies een pincode van minimaal 4 cijfers.'); return; }
      if (pin !== pin2) { setFout(`De ${wachtwoordLabel}s komen niet overeen.`); return; }
      setBezig(true);
      const gehasht = await hashWachtwoord(pin);
      setBezig(false);
      setPins({ ...pins, [naam]: gehasht });
      na2faControle();
    } else {
      setBezig(true);
      const ok = await verifieerWachtwoord(pin, pins[naam]);
      setBezig(false);
      if (ok) {
        // Soepele overgang: was dit nog een oude, niet-gehashte pincode? Dan nu alsnog hashen.
        if (typeof pins[naam] === 'string') {
          const gehasht = await hashWachtwoord(pin);
          setPins({ ...pins, [naam]: gehasht });
        }
        na2faControle();
      } else { setFout(`Onjuiste ${wachtwoordLabel}.`); }
    }
  }

  async function submitTfaSetup() {
    setFout('');
    setBezig(true);
    const ok = await verifieerTotp(nieuwSecret, tfaCode);
    setBezig(false);
    if (!ok) { setFout('Onjuiste code — controleer of de tijd op je telefoon klopt en probeer opnieuw.'); return; }
    setTfaSecrets({ ...tfaSecrets, [naam]: { secret: nieuwSecret, bevestigd: true } });
    markeerVertrouwd(naam);
    onLogin(gekozen);
  }
  async function submitTfaVerify() {
    setFout('');
    setBezig(true);
    const ok = await verifieerTotp(tfaSecrets[naam].secret, tfaCode);
    setBezig(false);
    if (!ok) { setFout('Onjuiste code.'); return; }
    markeerVertrouwd(naam);
    onLogin(gekozen);
  }

  if (stap === 'tfa-setup') {
    return (
      <div className="min-h-screen w-full flex items-center justify-center p-4" style={{ background: C.paper }}>
        <div className="w-full max-w-sm rounded-2xl border shadow-sm p-6" style={{ background: C.card, borderColor: C.border }}>
          <h2 className="text-center font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', fontSize: 18, color: C.ink }}>Twee-factor authenticatie instellen</h2>
          <p className="text-center text-xs mb-4" style={{ color: C.inkSoft }}>Eenmalig, voor {naam}</p>
          <div className="space-y-3">
            <p className="text-xs" style={{ color: C.inkSoft }}>Open een authenticator-app (bv. Google Authenticator of Microsoft Authenticator) en scan onderstaande QR-code, of voer de sleutel er handmatig in als scannen niet lukt:</p>
            <div className="flex justify-center py-1">
              <QrCode teksten={[
                `otpauth://totp/${encodeURIComponent(naam)}?secret=${nieuwSecret}&issuer=BladelsCreatief`,
                `otpauth://totp/${encodeURIComponent(naam)}?secret=${nieuwSecret}&issuer=Bladels`,
              ]} />
            </div>
            <p className="text-center font-mono text-sm rounded-lg px-3 py-2 tracking-widest" style={{ background: C.paperDim }}>{nieuwSecret}</p>
            <Field label="Voer de 6-cijferige code uit de app in ter bevestiging">
              <input inputMode="numeric" className={inputCls} style={inputStyle} value={tfaCode} onChange={e => setTfaCode(e.target.value.replace(/\D/g, '').slice(0, 6))} onKeyDown={e => e.key === 'Enter' && submitTfaSetup()} />
            </Field>
            {fout && <p className="text-xs" style={{ color: C.rose }}>{fout}</p>}
          </div>
          <div className="mt-4"><Btn onClick={submitTfaSetup} disabled={bezig || tfaCode.length !== 6}>{bezig ? 'Controleren…' : 'Bevestigen & aanmelden'}</Btn></div>
        </div>
      </div>
    );
  }
  if (stap === 'tfa-verify') {
    return (
      <div className="min-h-screen w-full flex items-center justify-center p-4" style={{ background: C.paper }}>
        <div className="w-full max-w-sm rounded-2xl border shadow-sm p-6" style={{ background: C.card, borderColor: C.border }}>
          <h2 className="text-center font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', fontSize: 18, color: C.ink }}>Verificatiecode</h2>
          <p className="text-center text-xs mb-4" style={{ color: C.inkSoft }}>Voer de code uit je authenticator-app in, {naam}</p>
          <div className="space-y-3">
            <Field label="6-cijferige code">
              <input inputMode="numeric" autoFocus className={inputCls} style={inputStyle} value={tfaCode} onChange={e => setTfaCode(e.target.value.replace(/\D/g, '').slice(0, 6))} onKeyDown={e => e.key === 'Enter' && submitTfaVerify()} />
            </Field>
            {fout && <p className="text-xs" style={{ color: C.rose }}>{fout}</p>}
          </div>
          <div className="mt-4"><Btn onClick={submitTfaVerify} disabled={bezig || tfaCode.length !== 6}>{bezig ? 'Controleren…' : 'Aanmelden'}</Btn></div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen w-full flex items-center justify-center p-4" style={{ background: C.paper }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');`}</style>
      <div className="w-full max-w-sm rounded-2xl border shadow-sm p-6" style={{ background: C.card, borderColor: C.border }}>
        <img src={LOGO_URI} alt="BladelsCreatief" style={{ height: 46, margin: '0 auto 14px', display: 'block' }} />
        <h2 className="text-center font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', fontSize: 18, color: C.ink }}>Aanmelden</h2>
        <p className="text-center text-xs mb-4" style={{ color: C.inkSoft }}>Alleen bestuursleden met een functie kunnen aanmelden.</p>

        <div className="space-y-3">
          <Field label="Ik ben">
            <select className={inputCls} style={inputStyle} value={naam} onChange={e => { setNaam(e.target.value); setPin(''); setPin2(''); setResetMode(false); setFout(''); }}>
              <option value="">— kies je naam —</option>
              {kandidaten.map(m => <option key={m.id} value={fullName(m)}>{fullName(m)} ({m.functie})</option>)}
            </select>
          </Field>

          {naam && (
            <>
              {isNieuw && (
                <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>
                  {resetMode ? `Kies een nieuwe ${wachtwoordLabel}.` : `Je meldt je voor het eerst aan — kies nu een ${wachtwoordLabel} voor jezelf.`}
                </p>
              )}
              <Field label={isNieuw ? `Nieuwe ${wachtwoordLabel}${beveiliging.wachtwoordEisen ? ' (min. 8 tekens, hoofdletter, kleine letter, cijfer, teken)' : ' (min. 4 cijfers)'}` : (wachtwoordLabel === 'wachtwoord' ? 'Wachtwoord' : 'Pincode')}>
                <input type="password" inputMode={beveiliging.wachtwoordEisen ? 'text' : 'numeric'} className={inputCls} style={inputStyle} value={pin}
                  onChange={e => setPin(beveiliging.wachtwoordEisen ? e.target.value : e.target.value.replace(/\D/g, ''))} onKeyDown={e => e.key === 'Enter' && !isNieuw && submitPincode()} />
              </Field>
              {isNieuw && (
                <Field label={`Herhaal ${wachtwoordLabel}`}>
                  <input type="password" inputMode={beveiliging.wachtwoordEisen ? 'text' : 'numeric'} className={inputCls} style={inputStyle} value={pin2}
                    onChange={e => setPin2(beveiliging.wachtwoordEisen ? e.target.value : e.target.value.replace(/\D/g, ''))} onKeyDown={e => e.key === 'Enter' && submitPincode()} />
                </Field>
              )}
              {heeftAlPin && !resetMode && (
                <button onClick={() => { setResetMode(true); setPin(''); }} className="text-xs underline" style={{ color: C.inkSoft }}>{wachtwoordLabel === 'wachtwoord' ? 'Wachtwoord' : 'Pincode'} vergeten? Nieuwe instellen</button>
              )}
            </>
          )}

          {fout && <p className="text-xs" style={{ color: C.rose }}>{fout}</p>}
        </div>

        <div className="mt-4">
          <Btn onClick={submitPincode} disabled={!naam || bezig}>{bezig ? 'Even controleren…' : (isNieuw ? `${wachtwoordLabel === 'wachtwoord' ? 'Wachtwoord' : 'Pincode'} instellen & aanmelden` : 'Aanmelden')}</Btn>
        </div>
        <p className="text-center text-xs mt-4" style={{ color: C.inkSoft, opacity: 0.6 }}>versie {APP_VERSIE}</p>
      </div>
    </div>
  );
}

function PromptModal({ title, label, placeholder, onSave, onCancel }) {
  const [value, setValue] = useState('');
  return (
    <Modal title={title} onClose={onCancel}>
      <Field label={label}>
        <input autoFocus className={inputCls} style={inputStyle} placeholder={placeholder} value={value}
          onChange={e => setValue(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter' && value.trim()) onSave(value.trim()); }} />
      </Field>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onCancel}>Annuleren</Btn>
        <Btn onClick={() => value.trim() && onSave(value.trim())} disabled={!value.trim()}>Toevoegen</Btn>
      </div>
    </Modal>
  );
}

function EmptyState({ icon: Icon, text }) {
  return (
    <div className="flex flex-col items-center justify-center py-12 text-center" style={{ color: C.inkSoft }}>
      <Icon size={28} className="mb-2 opacity-60" />
      <p className="text-sm">{text}</p>
    </div>
  );
}

/* ------------------------------ NAV -------------------------------- */
const TABS = [
  { id: 'dashboard', label: 'Overzicht', icon: LayoutDashboard },
  { id: 'leden', label: 'Leden', icon: Users },
  { id: 'workshops', label: 'Workshops', icon: Palette },
  { id: 'vergaderingen', label: 'Vergaderingen', icon: ClipboardList },
  { id: 'financien', label: 'Financiën', icon: Wallet },
  { id: 'begroting', label: 'Begroting', icon: PiggyBank },
  { id: 'rapportage', label: 'Rapportage', icon: FileSpreadsheet },
  { id: 'instellingen', label: 'Instellingen', icon: Settings },
];

/* =========================================================================
   APP
========================================================================= */
export default function BladelsCreatiefApp() {
  const [tab, setTab] = useState('dashboard');
  const [toast, setToast] = useState(null);
  const flash = (ok) => {
    setToast(ok ? { ok: true, tekst: 'Opgeslagen' } : { ok: false, tekst: 'Opslaan mislukt — controleer je verbinding' });
    setTimeout(() => setToast(null), ok ? 1400 : 3000);
  };

  const [members, setMembers] = useStored('bladels:leden', SEED_MEMBERS, true, flash);
  const [workshops, setWorkshops] = useStored('bladels:workshops', SEED_WORKSHOPS, true, flash);
  const [inschrijvingen, setInschrijvingen] = useStored('bladels:workshopinschrijvingen', SEED_WORKSHOP_INSCHRIJVINGEN, true, flash);
  const [tx, setTx] = useStored('bladels:transacties', SEED_TX, true, flash);
  const [accounts, setAccounts] = useStored('bladels:rekeningen', SEED_ACCOUNTS, true, flash);
  const [begrotingKoppelingen, setBegrotingKoppelingen] = useStored('bladels:begrotingkoppelingen', SEED_BEGROTING_KOPPELINGEN, true, flash);
  const [dagdelen, setDagdelen] = useStored('bladels:dagdelen', SEED_SLOTS, true, flash);
  const [workshopSoorten, setWorkshopSoorten] = useStored('bladels:workshopsoorten', SEED_WORKSHOP_SOORTEN, true, flash);
  const [agendapuntenVooraf, setAgendapuntenVooraf] = useStored('bladels:agendapuntenvooraf', SEED_AGENDAPUNTEN_VOORAF, true, flash);
  const [agendapuntenAfsluitend, setAgendapuntenAfsluitend] = useStored('bladels:agendapuntenafsluitend', SEED_AGENDAPUNTEN_AFSLUITEND, true, flash);
  useEffect(() => {
    const ontbrekend = SEED_ACCOUNTS.filter(a => !accounts.some(x => x.code === a.code));
    if (ontbrekend.length) setAccounts([...accounts, ...ontbrekend]);
  }, [accounts]);
  const [budget, setBudget] = useStored('bladels:begroting', SEED_BUDGET, true, flash);
  const [boekjaren, setBoekjaren] = useStored('bladels:boekjaren', SEED_YEARS, true, flash);
  const [vergaderingen, setVergaderingen] = useStored('bladels:vergaderingen', [], true, flash);
  const [actielijst, setActielijst] = useStored('bladels:actielijst', [], true, flash);
  const [contributies, setContributies] = useStored('bladels:contributies', [], true, flash);
  const [pins, setPins] = useStored('bladels:pins', {}, true, flash);
  const [beveiliging, setBeveiliging] = useStored('bladels:beveiliging', { tfaVerplicht: false, tfaVertrouwensdagen: 30, wachtwoordEisen: false }, true, flash);
  const [standaarden, setStandaarden] = useStored('bladels:standaarden', { idleTimeoutMinuten: 3, logoHoogteCm: 2.5 }, true, flash);
  const [workshopSortering, setWorkshopSortering] = useStored('bladels:workshopsortering', { veld: 'datum', richting: 'desc' }, true);
  const [tfaSecrets, setTfaSecrets] = useStored('bladels:tfa-secrets', {}, true);
  const [tfaVertrouwd, setTfaVertrouwd] = useStored('bladels:tfa-vertrouwd', {}, true);
  const [rolpermissies, setRolpermissies] = useStored('bladels:rolpermissies', SEED_ROLPERMISSIES, true, flash);
  const [logboek, setLogboek] = useStored('bladels:logboek', [], true);
  const [prullenbak, setPrullenbak] = useStored('bladels:prullenbak', [], true, flash);
  const [ingelogd, setIngelogd] = useState(null);
  const [pendingQuery, setPendingQuery] = useState('');

  async function stuurSessieSignaal(naam, verwijderen) {
    try {
      const res = await opslagLezen(STORAGE_KEYS.sessies, true);
      const huidig = res && res.value ? JSON.parse(res.value) : {};
      if (verwijderen) delete huidig[naam];
      else huidig[naam] = new Date().toISOString();
      await opslagSchrijven(STORAGE_KEYS.sessies, JSON.stringify(huidig), true);
    } catch (e) { /* best effort — presentie-indicator, geen kritieke data */ }
  }
  function handleLogin(persoon) {
    setIngelogd(persoon);
    setLogboek([{ id: uid(logboek), tijdstip: new Date().toISOString(), gebruiker: fullName(persoon), gebied: 'systeem', actie: 'Ingelogd' }, ...logboek].slice(0, 300));
    stuurSessieSignaal(fullName(persoon), false);
  }
  function handleLogout(reden) {
    if (ingelogd) {
      const actie = reden === 'inactiviteit' ? 'Automatisch uitgelogd wegens inactiviteit (3 minuten)' : 'Uitgelogd';
      setLogboek([{ id: uid(logboek), tijdstip: new Date().toISOString(), gebruiker: fullName(ingelogd), gebied: 'systeem', actie }, ...logboek].slice(0, 300));
      stuurSessieSignaal(fullName(ingelogd), true);
    }
    setIngelogd(null);
  }

  // Stuurt elke 45 seconden een "hartslag" naar de gedeelde opslag, zodat andere ingelogde
  // gebruikers (bv. vóór het herstellen van een back-up) kunnen zien wie er nog meer actief is.
  useEffect(() => {
    if (!ingelogd) return;
    const interval = setInterval(() => stuurSessieSignaal(fullName(ingelogd), false), 45000);
    return () => clearInterval(interval);
  }, [ingelogd]);

  // Automatisch uitloggen na een instelbaar aantal minuten zonder muis-/toetsenbord-/touch-activiteit.
  useEffect(() => {
    if (!ingelogd) return;
    let timer;
    function resetTimer() {
      clearTimeout(timer);
      timer = setTimeout(() => handleLogout('inactiviteit'), standaarden.idleTimeoutMinuten * 60 * 1000);
    }
    const events = ['mousemove', 'keydown', 'click', 'touchstart', 'scroll'];
    events.forEach(e => window.addEventListener(e, resetTimer));
    resetTimer();
    return () => {
      clearTimeout(timer);
      events.forEach(e => window.removeEventListener(e, resetTimer));
    };
  }, [ingelogd, standaarden.idleTimeoutMinuten]);

  if (!ingelogd) {
    return <LoginScreen members={members} pins={pins} setPins={setPins} beveiliging={beveiliging}
      tfaSecrets={tfaSecrets} setTfaSecrets={setTfaSecrets} tfaVertrouwd={tfaVertrouwd} setTfaVertrouwd={setTfaVertrouwd}
      onLogin={handleLogin} />;
  }

  const magBewerken = new Set(bewerkbareTabs(ingelogd.functie, rolpermissies));
  const readOnly = tabId => !magBewerken.has(tabId);
  const isVoorzitter = /voorzitter/i.test(ingelogd.functie || '');

  function goTo(tabId, query) {
    setPendingQuery(query || '');
    setTab(tabId);
  }
  function logAction(actie, gebied) {
    const entry = { id: uid(logboek), tijdstip: new Date().toISOString(), gebruiker: fullName(ingelogd), gebied, actie };
    setLogboek([entry, ...logboek].slice(0, 300));
  }
  function trashIt(type, data) {
    const entry = { id: uid(prullenbak), type, data, verwijderdOp: new Date().toISOString(), verwijderdDoor: fullName(ingelogd) };
    setPrullenbak([entry, ...prullenbak]);
  }

  return (
    <div className="min-h-screen w-full" style={{ background: C.paper, color: C.ink, fontFamily: 'Inter, ui-sans-serif, system-ui' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        ::selection { background: ${C.clay}33; }
        table { border-collapse: collapse; }
        input:focus, select:focus, textarea:focus { box-shadow: 0 0 0 3px ${C.clay}22; border-color: ${C.clay}; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 8px; }
      `}</style>

      <Header tab={tab} setTab={setTab} ingelogd={ingelogd} onLogout={() => handleLogout('handmatig')}
        members={members} workshops={workshops} inschrijvingen={inschrijvingen} tx={tx} vergaderingen={vergaderingen} actielijst={actielijst} onNavigate={goTo} readOnly={readOnly} />

      <main className="max-w-7xl mx-auto px-4 sm:px-6 pb-16 pt-5">
        {tab === 'dashboard' && <Dashboard members={members} workshops={workshops} inschrijvingen={inschrijvingen} tx={tx} budget={budget} setTab={setTab} dagdelen={dagdelen} />}
        {tab === 'leden' && <LedenTab members={members} setMembers={setMembers} readOnly={readOnly('leden')}
          contributies={contributies} setContributies={setContributies} boekjaren={boekjaren} tx={tx} setTx={setTx} accounts={accounts}
          dagdelen={dagdelen}
          initialQuery={pendingQuery} onTrash={trashIt} onLog={logAction} />}
        {tab === 'workshops' && <WorkshopsTab members={members} workshops={workshops} setWorkshops={setWorkshops}
          inschrijvingen={inschrijvingen} setInschrijvingen={setInschrijvingen} tx={tx} setTx={setTx} accounts={accounts}
          workshopSoorten={workshopSoorten} logoHoogteCm={standaarden.logoHoogteCm}
          workshopSortering={workshopSortering} setWorkshopSortering={setWorkshopSortering}
          readOnly={readOnly('workshops')} initialQuery={pendingQuery} onTrash={trashIt} onLog={logAction} />}
        {tab === 'vergaderingen' && <VergaderingenTab members={members} vergaderingen={vergaderingen} setVergaderingen={setVergaderingen} actielijst={actielijst} setActielijst={setActielijst}
          agendapuntenVooraf={agendapuntenVooraf} agendapuntenAfsluitend={agendapuntenAfsluitend} logoHoogteCm={standaarden.logoHoogteCm}
          readOnly={readOnly('vergaderingen')} onTrash={trashIt} onLog={logAction} />}
        {tab === 'financien' && <FinancienTab tx={tx} setTx={setTx} accounts={accounts} boekjaren={boekjaren} setBoekjaren={setBoekjaren}
          members={members} contributies={contributies} setContributies={setContributies}
          workshops={workshops} inschrijvingen={inschrijvingen} setInschrijvingen={setInschrijvingen}
          readOnly={readOnly('financien')} initialQuery={pendingQuery} onTrash={trashIt} onLog={logAction} />}
        {tab === 'begroting' && <BegrotingTab budget={budget} setBudget={setBudget} tx={tx} boekjaren={boekjaren} setBoekjaren={setBoekjaren} begrotingKoppelingen={begrotingKoppelingen} readOnly={readOnly('begroting')} onLog={logAction} />}
        {tab === 'rapportage' && <RapportageTab tx={tx} accounts={accounts} members={members} workshops={workshops} inschrijvingen={inschrijvingen} budget={budget} begrotingKoppelingen={begrotingKoppelingen} logoHoogteCm={standaarden.logoHoogteCm} />}
        {tab === 'instellingen' && (
          <InstellingenTab isVoorzitter={isVoorzitter} ingelogd={ingelogd} rolpermissies={rolpermissies} setRolpermissies={setRolpermissies}
            beveiliging={beveiliging} setBeveiliging={setBeveiliging} standaarden={standaarden} setStandaarden={setStandaarden}
            logboek={logboek} prullenbak={prullenbak} setPrullenbak={setPrullenbak}
            members={members} setMembers={setMembers} workshops={workshops} setWorkshops={setWorkshops}
            inschrijvingen={inschrijvingen} setInschrijvingen={setInschrijvingen}
            tx={tx} setTx={setTx} boekjaren={boekjaren} setBoekjaren={setBoekjaren} accounts={accounts} setAccounts={setAccounts}
            begrotingKoppelingen={begrotingKoppelingen} setBegrotingKoppelingen={setBegrotingKoppelingen}
            dagdelen={dagdelen} setDagdelen={setDagdelen}
            workshopSoorten={workshopSoorten} setWorkshopSoorten={setWorkshopSoorten}
            agendapuntenVooraf={agendapuntenVooraf} setAgendapuntenVooraf={setAgendapuntenVooraf}
            agendapuntenAfsluitend={agendapuntenAfsluitend} setAgendapuntenAfsluitend={setAgendapuntenAfsluitend}
            budget={budget} setBudget={setBudget} actielijst={actielijst} setActielijst={setActielijst}
            contributies={contributies} setContributies={setContributies}
            vergaderingen={vergaderingen} setVergaderingen={setVergaderingen}
            magLeden={magBewerken.has('leden')} magWorkshops={magBewerken.has('workshops')} magFinancien={magBewerken.has('financien')} magVergaderingen={magBewerken.has('vergaderingen')}
            magLedenImporteren={magBewerken.has('leden')} magWorkshopsImporteren={magBewerken.has('workshops')} magFinancienImporteren={magBewerken.has('financien')}
            backupData={{ members, workshops, inschrijvingen, tx, accounts, budget, boekjaren, vergaderingen, actielijst, contributies, rolpermissies, begrotingKoppelingen }}
            onLog={logAction} />
        )}
      </main>

      {toast && (
        <div className="fixed bottom-4 right-4 z-50 flex items-center gap-2 px-3.5 py-2 rounded-lg shadow-lg text-sm"
          style={{ background: toast.ok ? C.sageDeep : C.rose, color: '#fff' }}>
          {toast.ok ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />} {toast.tekst}
        </div>
      )}
    </div>
  );
}

function Header({ tab, setTab, ingelogd, onLogout, members, workshops, inschrijvingen, tx, vergaderingen, actielijst, onNavigate, readOnly }) {
  const editeerbareIds = new Set(EDITEERBARE_TABS.map(t => t.id));
  const geordendeTabs = [...TABS].sort((a, b) => {
    const aGeblokkeerd = editeerbareIds.has(a.id) && readOnly(a.id);
    const bGeblokkeerd = editeerbareIds.has(b.id) && readOnly(b.id);
    if (aGeblokkeerd === bGeblokkeerd) return 0;
    return aGeblokkeerd ? 1 : -1;
  });
  return (
    <header className="sticky top-0 z-40 border-b backdrop-blur"
      style={{ background: `${C.paper}F2`, borderColor: C.border }}>
      <div className="max-w-7xl mx-auto px-4 sm:px-6 pt-4 pb-3">
        <div className="flex items-center gap-2.5 mb-3 justify-between flex-wrap">
          <div className="flex items-center gap-2.5">
            <img src={LOGO_URI} alt="BladelsCreatief" style={{ height: 44, width: 'auto', display: 'block' }} />
            <div className="border-l pl-2.5" style={{ borderColor: C.border }}>
              <p className="text-xs font-medium" style={{ color: C.ink }}>Leden- &amp; financiële administratie</p>
            </div>
          </div>
          <div className="flex items-center gap-3 flex-wrap">
            {ingelogd && <GlobalSearch members={members} workshops={workshops} inschrijvingen={inschrijvingen} tx={tx} vergaderingen={vergaderingen} actielijst={actielijst} onNavigate={onNavigate} />}
            {ingelogd && (
              <div className="flex items-center gap-2 text-xs" style={{ color: C.inkSoft }}>
                <span>Ingelogd als <strong style={{ color: C.ink }}>{fullName(ingelogd)}</strong> ({ingelogd.functie})</span>
                <button onClick={onLogout} className="underline" style={{ color: C.clay }}>Uitloggen</button>
              </div>
            )}
          </div>
        </div>
        <nav className="flex gap-1 overflow-x-auto -mb-px">
          {geordendeTabs.map(t => {
            const active = tab === t.id;
            const Icon = t.icon;
            const geblokkeerd = editeerbareIds.has(t.id) && readOnly(t.id);
            return (
              <button key={t.id} onClick={() => setTab(t.id)}
                className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium rounded-t-lg border-b-2 whitespace-nowrap transition-colors"
                style={{
                  borderColor: active ? C.clay : 'transparent',
                  color: active ? C.clay : (geblokkeerd ? '#A7B2C4' : C.inkSoft),
                  background: active ? C.card : 'transparent',
                  opacity: geblokkeerd ? 0.7 : 1,
                }}>
                <Icon size={15} /> {t.label}
              </button>
            );
          })}
        </nav>
      </div>
    </header>
  );
}

function GlobalSearch({ members, workshops, inschrijvingen, tx, vergaderingen, actielijst, onNavigate }) {
  const [q, setQ] = useState('');
  const [open, setOpen] = useState(false);
  const term = q.trim().toLowerCase();

  const ledenHits = term ? members.filter(m => `${fullName(m)} ${m.email || ''} ${m.telefoon || ''} ${m.woonplaats || ''}`.toLowerCase().includes(term)).slice(0, 5) : [];
  const workshopHits = term ? inschrijvingen.filter(i => (i.naam || '').toLowerCase().includes(term)).slice(0, 5) : [];
  const txHits = term ? tx.filter(t => `${t.omschrijving || ''} ${t.grootboek_naam || ''}`.toLowerCase().includes(term)).slice(0, 5) : [];
  const vergHits = term ? vergaderingen.filter(v => (v.titel || '').toLowerCase().includes(term)).slice(0, 5) : [];
  const actieHits = term ? actielijst.filter(a => (a.omschrijving || '').toLowerCase().includes(term)).slice(0, 5) : [];
  const totaal = ledenHits.length + workshopHits.length + txHits.length + vergHits.length + actieHits.length;

  function ga(tabId, query) {
    onNavigate(tabId, query);
    setOpen(false);
    setQ('');
  }

  return (
    <div className="relative">
      <div className="relative">
        <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: C.inkSoft }} />
        <input value={q} onFocus={() => setOpen(true)} onChange={e => { setQ(e.target.value); setOpen(true); }}
          placeholder="Zoek overal…" className={`${inputCls} pl-8`} style={{ ...inputStyle, width: 210 }} />
      </div>
      {open && term && (
        <div className="absolute right-0 mt-1 w-80 max-h-96 overflow-y-auto rounded-xl border shadow-lg z-50" style={{ background: C.card, borderColor: C.border }}
          onMouseLeave={() => setOpen(false)}>
          {totaal === 0 && <p className="text-xs px-3 py-3" style={{ color: C.inkSoft }}>Geen resultaten voor "{q}".</p>}
          {ledenHits.length > 0 && (
            <div className="p-2 border-b" style={{ borderColor: C.border }}>
              <p className="text-xs font-medium px-1 mb-1" style={{ color: C.inkSoft }}>Leden</p>
              {ledenHits.map(m => (
                <button key={m.id} onClick={() => ga('leden', fullName(m))} className="w-full text-left text-sm px-2 py-1.5 rounded hover:bg-black/5">{fullName(m)}</button>
              ))}
            </div>
          )}
          {workshopHits.length > 0 && (
            <div className="p-2 border-b" style={{ borderColor: C.border }}>
              <p className="text-xs font-medium px-1 mb-1" style={{ color: C.inkSoft }}>Workshops</p>
              {workshopHits.map(i => (
                <button key={i.id} onClick={() => ga('workshops', i.naam)} className="w-full text-left text-sm px-2 py-1.5 rounded hover:bg-black/5">{i.naam}</button>
              ))}
            </div>
          )}
          {txHits.length > 0 && (
            <div className="p-2 border-b" style={{ borderColor: C.border }}>
              <p className="text-xs font-medium px-1 mb-1" style={{ color: C.inkSoft }}>Financiën</p>
              {txHits.map(t => (
                <button key={t.id} onClick={() => ga('financien', t.omschrijving)} className="w-full text-left text-sm px-2 py-1.5 rounded hover:bg-black/5">
                  {t.omschrijving || t.grootboek_naam} <span style={{ color: C.inkSoft }}>· {euro(t.bedrag)}</span>
                </button>
              ))}
            </div>
          )}
          {vergHits.length > 0 && (
            <div className="p-2 border-b" style={{ borderColor: C.border }}>
              <p className="text-xs font-medium px-1 mb-1" style={{ color: C.inkSoft }}>Vergaderingen</p>
              {vergHits.map(v => (
                <button key={v.id} onClick={() => ga('vergaderingen', '')} className="w-full text-left text-sm px-2 py-1.5 rounded hover:bg-black/5">{v.titel} <span style={{ color: C.inkSoft }}>· {fmtDate(v.datum)}</span></button>
              ))}
            </div>
          )}
          {actieHits.length > 0 && (
            <div className="p-2">
              <p className="text-xs font-medium px-1 mb-1" style={{ color: C.inkSoft }}>Actiepunten</p>
              {actieHits.map(a => (
                <button key={a.id} onClick={() => ga('vergaderingen', '')} className="w-full text-left text-sm px-2 py-1.5 rounded hover:bg-black/5">{a.omschrijving}</button>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* =========================================================================
   DASHBOARD
========================================================================= */
function Dashboard({ members, workshops, inschrijvingen, tx, budget, setTab, dagdelen }) {
  const actief = members.filter(m => m.status === 'actief');
  const inactief = members.length - actief.length;
  const actieveInschrijvingen = inschrijvingen.filter(i => i.status === 'ingeschreven');
  const workshopDeelnemers = new Set(actieveInschrijvingen.map(i => i.naam)).size;
  const wachtlijst = inschrijvingen.filter(i => i.status === 'wachtlijst').length;

  const bal908 = START_BALANCE['908'] + tx.filter(t => t.rekening === '908').reduce((s, t) => s + Number(t.bedrag || 0), 0);
  const bal319 = START_BALANCE['319'] + tx.filter(t => t.rekening === '319').reduce((s, t) => s + Number(t.bedrag || 0), 0);
  const saldoTotaal = bal908 + bal319;

  const years = Array.from(new Set(tx.map(t => t.jaar))).sort();
  const huidigJaar = years.length ? years[years.length - 1] : new Date().getFullYear();
  const txJaar = tx.filter(t => t.jaar === huidigJaar);
  const inkomsten = txJaar.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0);
  const uitgaven = txJaar.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0);

  const bJaar = budget[String(huidigJaar)] || [];
  const begrootInkomsten = bJaar.filter(b => b.sectie === 'Inkomsten').reduce((s, b) => s + Number(b.bedrag || 0), 0);
  const begrootUitgaven = bJaar.filter(b => b.sectie === 'Uitgaven').reduce((s, b) => s + Number(b.bedrag || 0), 0);

  const monthly = MONTH_NAMES.map((naam, i) => {
    const maand = i + 1;
    const rows = txJaar.filter(t => Number(t.maand) === maand);
    return {
      naam,
      Inkomsten: Math.round(rows.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0)),
      Uitgaven: Math.round(-rows.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0)),
    };
  });

  const groepen = dagdelen.map(slot => ({
    slot, count: actief.filter(m => (m.dagdelen || []).includes(slot)).length
  }));

  const recent = [...tx].sort((a, b) => (b.datum || '').localeCompare(a.datum || '')).slice(0, 6);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <KpiCard icon={Users} label="Actieve leden" value={actief.length} sub={`${inactief} inactief`} tone="clay" onClick={() => setTab('leden')} />
        <KpiCard icon={Palette} label="Workshops" value={workshopDeelnemers} sub={`unieke deelnemers${wachtlijst ? ` · ${wachtlijst} op wachtlijst` : ''}`} tone="ochre" onClick={() => setTab('workshops')} />
        <KpiCard icon={Landmark} label="Totaal saldo" value={euro(saldoTotaal)} sub={`.908: ${euro(bal908)} · .319: ${euro(bal319)}`} tone="sage" onClick={() => setTab('financien')} />
        <KpiCard icon={TrendingUp} label={`Resultaat ${huidigJaar}`} value={euro(inkomsten + uitgaven)} sub={`${euro(inkomsten)} in · ${euro(uitgaven)} uit`} tone={inkomsten + uitgaven >= 0 ? 'sage' : 'rose'} onClick={() => setTab('rapportage')} />
      </div>

      <div className="grid lg:grid-cols-3 gap-4">
        <Card className="lg:col-span-2 p-4">
          <div className="flex items-center justify-between mb-3">
            <h3 className="font-semibold" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Inkomsten &amp; uitgaven per maand · {huidigJaar}</h3>
          </div>
          <div style={{ height: 260 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={monthly} barGap={3}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.border} vertical={false} />
                <XAxis dataKey="naam" tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={false} tickLine={false} width={40} />
                <Tooltip formatter={(v) => euro(v)} contentStyle={{ borderRadius: 8, borderColor: C.border, fontSize: 12 }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Bar dataKey="Inkomsten" fill={C.sage} radius={[3, 3, 0, 0]} />
                <Bar dataKey="Uitgaven" fill={C.clay} radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>

        <Card className="p-4">
          <h3 className="font-semibold mb-3" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Begroting {huidigJaar}</h3>
          <BudgetBar label="Inkomsten" begroot={begrootInkomsten} werkelijk={inkomsten} tone="sage" />
          <BudgetBar label="Uitgaven" begroot={begrootUitgaven} werkelijk={-uitgaven} tone="clay" />
          <button onClick={() => setTab('begroting')} className="text-xs mt-2 underline" style={{ color: C.clay }}>Bekijk volledige begroting →</button>
        </Card>
      </div>

      <div className="grid lg:grid-cols-3 gap-4">
        <Card className="lg:col-span-2 p-4">
          <div className="flex items-center gap-2 mb-3">
            <Grid3x3 size={16} style={{ color: C.clay }} />
            <h3 className="font-semibold" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Groepsindeling (actieve leden per dagdeel)</h3>
          </div>
          <div className="grid sm:grid-cols-2 gap-2">
            {groepen.map(g => (
              <div key={g.slot} className="flex items-center justify-between px-3 py-2 rounded-lg" style={{ background: C.paperDim }}>
                <span className="text-sm">{g.slot}</span>
                <Badge tone="clay">{g.count} leden</Badge>
              </div>
            ))}
          </div>
        </Card>

        <Card className="p-4">
          <h3 className="font-semibold mb-3" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Recente boekingen</h3>
          <div className="space-y-2">
            {recent.map(t => (
              <div key={t.id} className="flex items-center justify-between text-sm">
                <div className="min-w-0 pr-2">
                  <p className="truncate">{t.omschrijving || t.grootboek_naam}</p>
                  <p className="text-xs" style={{ color: C.inkSoft }}>{fmtDate(t.datum)}</p>
                </div>
                <span className="font-medium whitespace-nowrap" style={{ color: t.bedrag >= 0 ? C.sageDeep : C.rose }}>
                  {t.bedrag >= 0 ? '+' : ''}{euro(t.bedrag)}
                </span>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

function KpiCard({ icon: Icon, label, value, sub, tone, onClick }) {
  const bg = { clay: C.clay, sage: C.sage, ochre: C.ochre, rose: C.rose }[tone] || C.clay;
  return (
    <button onClick={onClick} className="text-left p-4 rounded-xl border shadow-sm hover:shadow-md transition-shadow"
      style={{ background: C.card, borderColor: C.border }}>
      <div className="w-8 h-8 rounded-lg grid place-items-center mb-2" style={{ background: `${bg}1F` }}>
        <Icon size={16} style={{ color: bg }} />
      </div>
      <p className="text-xs" style={{ color: C.inkSoft }}>{label}</p>
      <p className="text-xl font-semibold" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>{value}</p>
      {sub && <p className="text-xs mt-0.5" style={{ color: C.inkSoft }}>{sub}</p>}
    </button>
  );
}

function BudgetBar({ label, begroot, werkelijk, tone }) {
  const pct = begroot > 0 ? Math.min(100, Math.round((werkelijk / begroot) * 100)) : 0;
  const color = tone === 'clay' ? C.clay : C.sage;
  return (
    <div className="mb-3">
      <div className="flex justify-between text-xs mb-1">
        <span style={{ color: C.inkSoft }}>{label}</span>
        <span style={{ color: C.ink }}>{euro(werkelijk)} / {euro(begroot)}</span>
      </div>
      <div className="h-2 rounded-full overflow-hidden" style={{ background: C.paperDim }}>
        <div className="h-full rounded-full" style={{ width: `${pct}%`, background: color }} />
      </div>
    </div>
  );
}

/* =========================================================================
   LEDEN
========================================================================= */
function LedenTab({ members, setMembers, readOnly, contributies, setContributies, boekjaren, tx, setTx, accounts, dagdelen, initialQuery, onTrash, onLog }) {
  const [q, setQ] = useState(initialQuery || '');
  const [status, setStatus] = useState(initialQuery ? 'alle' : 'actief');
  const [dagdeel, setDagdeel] = useState('alle');
  const [view, setView] = useState('lijst');
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [showReports, setShowReports] = useState(false);
  const [waSelectie, setWaSelectie] = useState([]);
  const [showBulkWa, setShowBulkWa] = useState(false);
  function toggleWaSelectie(id) {
    setWaSelectie(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
  }

  const filtered = members.filter(m => {
    if (status !== 'alle' && m.status !== status) return false;
    if (dagdeel !== 'alle' && !(m.dagdelen || []).includes(dagdeel)) return false;
    if (q) {
      const hay = `${fullName(m)} ${m.woonplaats || ''} ${m.email || ''} ${m.functie || ''}`.toLowerCase();
      if (!hay.includes(q.toLowerCase())) return false;
    }
    return true;
  }).sort((a, b) => fullName(a).localeCompare(fullName(b)));

  function saveMember(data) {
    if (data.id) {
      setMembers(members.map(m => m.id === data.id ? data : m));
      onLog(`Lid bewerkt: ${fullName(data)}`, 'leden');
    } else {
      setMembers([...members, { ...data, id: uid(members) }]);
      onLog(`Lid toegevoegd: ${fullName(data)}`, 'leden');
    }
    setShowForm(false); setEditing(null);
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-2 justify-between">
        <div className="flex flex-wrap items-center gap-2">
          <div className="relative">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: C.inkSoft }} />
            <input value={q} onChange={e => setQ(e.target.value)} placeholder="Zoek lid…"
              className={`${inputCls} pl-8`} style={{ ...inputStyle, width: 200 }} />
          </div>
          <select value={status} onChange={e => setStatus(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Alle statussen</option>
            <option value="actief">Actief</option>
            <option value="inactief">Inactief</option>
          </select>
          <select value={dagdeel} onChange={e => setDagdeel(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Alle dagdelen</option>
            {dagdelen.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <div className="flex rounded-lg border overflow-hidden" style={{ borderColor: C.border }}>
            <button onClick={() => setView('lijst')} className="px-2.5 py-1.5 text-xs flex items-center gap-1"
              style={{ background: view === 'lijst' ? C.paperDim : 'transparent', color: C.ink }}><ListChecks size={13} /> Lijst</button>
            <button onClick={() => setView('groepen')} className="px-2.5 py-1.5 text-xs flex items-center gap-1"
              style={{ background: view === 'groepen' ? C.paperDim : 'transparent', color: C.ink }}><Grid3x3 size={13} /> Groepen</button>
            <button onClick={() => setView('contributie')} className="px-2.5 py-1.5 text-xs flex items-center gap-1"
              style={{ background: view === 'contributie' ? C.paperDim : 'transparent', color: C.ink }}><PiggyBank size={13} /> Contributie</button>
          </div>
        </div>
        <div className="flex gap-2">
          <Btn icon={FileSpreadsheet} tone="outline" onClick={() => setShowReports(true)}>Overzicht maken</Btn>
          {!readOnly && <Btn icon={Plus} onClick={() => { setEditing(null); setShowForm(true); }}>Nieuw lid</Btn>}
        </div>
      </div>

      <p className="text-xs" style={{ color: C.inkSoft }}>{filtered.length} van {members.length} leden</p>

      {view === 'lijst' && (
        <Card className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left border-b" style={{ borderColor: C.border }}>
                {['Naam', 'Woonplaats', 'Telefoon', 'Functie', 'Dagdelen', 'Groepsapp', 'Status', 'Lid sinds', 'Lidmaatschap beëindigd', ''].map(h => (
                  <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(m => (
                <tr key={m.id} className="border-b last:border-0 hover:bg-black/[0.02]" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2 font-medium">{fullName(m)}</td>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{m.woonplaats || '—'}</td>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{m.telefoon || '—'}</td>
                  <td className="px-3 py-2">{m.functie ? <Badge tone="ochre">{m.functie}</Badge> : <span style={{ color: C.inkSoft }}>—</span>}</td>
                  <td className="px-3 py-2">
                    <div className="flex flex-wrap gap-1">
                      {(m.dagdelen || []).map(d => <Badge key={d} tone="sage">{d.split(' ')[0]}</Badge>)}
                      {(!m.dagdelen || !m.dagdelen.length) && <span style={{ color: C.inkSoft }}>—</span>}
                    </div>
                  </td>
                  <td className="px-3 py-2">
                    {(!m.groepsapp || m.groepsapp === 'ja')
                      ? <span style={{ color: C.inkSoft }}>Ja</span>
                      : <Badge tone={m.groepsapp === 'afgemeld' ? 'rose' : 'muted'}>{ledenVeldWaarde(m, 'groepsapp')}</Badge>}
                  </td>
                  <td className="px-3 py-2">
                    <Badge tone={m.status === 'actief' ? 'sage' : 'muted'}>{m.status || '—'}</Badge>
                  </td>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{fmtDate(m.lidsinds)}</td>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{m.status === 'inactief' && m.eindelidmaat ? fmtDate(m.eindelidmaat) : '—'}</td>
                  <td className="px-3 py-2">
                    <div className="flex gap-1 justify-end">
                      {!readOnly && (
                        <button onClick={() => { setEditing(m); setShowForm(true); }} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={14} /></button>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {!filtered.length && <EmptyState icon={Users} text="Geen leden gevonden met deze filters." />}
        </Card>
      )}

      {view === 'groepen' && (
        <div className="space-y-3">
          <div className="grid sm:grid-cols-2 gap-3">
            {dagdelen.map(slot => {
              const leden = members.filter(m => m.status === 'actief' && (m.dagdelen || []).includes(slot))
                .sort((a, b) => fullName(a).localeCompare(fullName(b)));
              const alleGeselecteerd = leden.length > 0 && leden.every(m => waSelectie.includes(m.id));
              function toggleGroep() {
                if (alleGeselecteerd) setWaSelectie(s => s.filter(id => !leden.some(m => m.id === id)));
                else setWaSelectie(s => Array.from(new Set([...s, ...leden.map(m => m.id)])));
              }
              return (
                <Card key={slot} className="p-4">
                  <div className="flex items-center justify-between mb-2">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input type="checkbox" checked={alleGeselecteerd} onChange={toggleGroep} disabled={!leden.length} />
                      <h4 className="font-semibold text-sm">{slot}</h4>
                    </label>
                    <Badge tone="clay">{leden.length}</Badge>
                  </div>
                  <ul className="text-sm space-y-1.5">
                    {leden.map(m => (
                      <li key={m.id} className="flex items-center gap-2">
                        <input type="checkbox" checked={waSelectie.includes(m.id)} onChange={() => toggleWaSelectie(m.id)} />
                        <span className="flex-1" style={{ color: C.ink }}>{fullName(m)}</span>
                        {m.telefoon && <span className="text-xs" style={{ color: C.inkSoft, userSelect: 'text', WebkitUserSelect: 'text' }} title="Tik en houd vast om te selecteren en kopiëren">{m.telefoon}</span>}
                        <ContactActies telefoon={m.telefoon} naam={fullName(m)} />
                      </li>
                    ))}
                    {!leden.length && <li style={{ color: C.inkSoft }} className="italic">Geen leden in dit dagdeel</li>}
                  </ul>
                </Card>
              );
            })}
          </div>

          {waSelectie.length > 0 && (
            <div className="sticky bottom-3 flex items-center justify-between gap-3 px-4 py-2.5 rounded-xl shadow-lg" style={{ background: C.ink, color: '#fff' }}>
              <span className="text-sm">{waSelectie.length} lid/leden geselecteerd</span>
              <div className="flex gap-2">
                <button onClick={() => setWaSelectie([])} className="text-xs underline">wis selectie</button>
                <Btn tone="sage" size="sm" icon={MessageCircle} onClick={() => setShowBulkWa(true)}>WhatsApp-bericht</Btn>
              </div>
            </div>
          )}
        </div>
      )}

      {view === 'contributie' && (
        <ContributieView members={filtered} contributies={contributies} setContributies={setContributies}
          boekjaren={boekjaren} tx={tx} setTx={setTx} accounts={accounts} readOnly={readOnly} onLog={onLog} />
      )}

      {showBulkWa && (
        <BulkWhatsAppModal leden={members.filter(m => waSelectie.includes(m.id))} onClose={() => setShowBulkWa(false)} />
      )}

      {showForm && (
        <MemberForm member={editing} dagdelen={dagdelen} onSave={saveMember} onClose={() => { setShowForm(false); setEditing(null); }} />
      )}
      {showReports && (
        <ReportsModal members={members} dagdelen={dagdelen} onClose={() => setShowReports(false)} />
      )}
    </div>
  );
}

function ContributieView({ members, contributies, setContributies, boekjaren, tx, setTx, accounts, readOnly, onLog }) {
  const huidigJaar = new Date().getFullYear();
  const [jaar, setJaar] = useState(boekjaren.includes(huidigJaar) ? huidigJaar : (boekjaren.length ? boekjaren[boekjaren.length - 1] : huidigJaar));
  const [betaalFor, setBetaalFor] = useState(null);

  function recordFor(m) {
    return contributies.find(c => c.lidId === m.id && Number(c.jaar) === Number(jaar));
  }
  const betaaldCount = members.filter(m => { const r = recordFor(m); return r && r.betaald; }).length;

  function markeerBetaald(member, data) {
    const bestaand = recordFor(member);
    const patch = { bedrag: Number(data.bedrag), datum: data.datum, betaald: true };
    if (bestaand) {
      setContributies(contributies.map(c => c.id === bestaand.id ? { ...c, ...patch } : c));
    } else {
      setContributies([...contributies, { id: uid(contributies), lidId: member.id, jaar: Number(jaar), ...patch }]);
    }
    onLog(`Contributie ${jaar} gemarkeerd als betaald: ${fullName(member)} (${euro(patch.bedrag)})`, 'leden');
    if (data.boeken) {
      const acc = accounts.find(a => a.naam.toLowerCase().includes('contributie')) || { code: '8002', naam: 'Contributie bijdragen' };
      const nieuw = {
        id: uid(tx), rekening: data.rekening, jaar: Number(jaar), maand: new Date(data.datum).getMonth() + 1,
        datum: data.datum, grootboek_code: acc.code, grootboek_naam: acc.naam,
        bedrag: Math.abs(Number(data.bedrag)), omschrijving: `Contributie ${jaar} - ${fullName(member)}`,
        gekoppeldType: 'contributie', gekoppeldRef: { lidId: member.id, jaar: Number(jaar) },
      };
      setTx([...tx, nieuw]);
      onLog(`Boeking toegevoegd voor contributie: ${fullName(member)} (${euro(nieuw.bedrag)})`, 'financien');
    }
    setBetaalFor(null);
  }
  const [openZonderKoppeling, setOpenZonderKoppeling] = useState(null);
  function gekoppeldeBoeking(member) {
    return tx.find(t => t.gekoppeldType === 'contributie' && t.gekoppeldRef.lidId === member.id && String(t.gekoppeldRef.jaar) === String(jaar));
  }
  function markeerOpen(member) {
    const boeking = gekoppeldeBoeking(member);
    if (boeking) { setOpenZonderKoppeling(member); return; }
    doeMarkeerOpen(member);
  }
  function doeMarkeerOpen(member) {
    const bestaand = recordFor(member);
    if (bestaand) {
      setContributies(contributies.map(c => c.id === bestaand.id ? { ...c, betaald: false } : c));
      onLog(`Contributie ${jaar} gemarkeerd als open: ${fullName(member)}`, 'leden');
    }
    const boeking = gekoppeldeBoeking(member);
    if (boeking) setTx(tx.map(t => t.id === boeking.id ? { ...t, gekoppeldType: null, gekoppeldRef: null } : t));
    setOpenZonderKoppeling(null);
  }

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <select value={jaar} onChange={e => setJaar(e.target.value)} className={inputCls} style={inputStyle}>
          {(boekjaren.length ? boekjaren : [jaar]).map(j => <option key={j} value={j}>{j}</option>)}
        </select>
        <Badge tone={betaaldCount === members.length ? 'sage' : 'ochre'}>{betaaldCount} / {members.length} betaald</Badge>
      </div>
      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              {['Naam', 'Status', 'Bedrag', 'Datum ontvangen', ''].map(h => (
                <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {members.map(m => {
              const r = recordFor(m);
              const betaald = r && r.betaald;
              return (
                <tr key={m.id} className="border-b last:border-0" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2 font-medium">{fullName(m)}</td>
                  <td className="px-3 py-2"><Badge tone={betaald ? 'sage' : 'rose'}>{betaald ? 'Betaald' : 'Open'}</Badge></td>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{r ? euro(r.bedrag) : '—'}</td>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{r && r.datum ? fmtDate(r.datum) : '—'}</td>
                  <td className="px-3 py-2 text-right">
                    {!readOnly && (betaald
                      ? <button onClick={() => markeerOpen(m)} className="text-xs underline" style={{ color: C.inkSoft }}>markeer open</button>
                      : <button onClick={() => setBetaalFor(m)} className="text-xs underline" style={{ color: C.clay }}>markeer betaald</button>)}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {!members.length && <EmptyState icon={PiggyBank} text="Geen leden om contributie voor bij te houden (pas eventueel de filters hierboven aan)." />}
      </Card>
      {betaalFor && (
        <ContributieForm member={betaalFor} jaar={jaar} onSave={data => markeerBetaald(betaalFor, data)} onClose={() => setBetaalFor(null)} />
      )}
      {openZonderKoppeling && (
        <ConfirmModal title="Contributie op open zetten"
          message={`Dit is gekoppeld aan een boeking in Financiën (${gekoppeldeBoeking(openZonderKoppeling)?.omschrijving || ''}). Doorgaan zet de contributie op open en ontkoppelt de boeking — de boeking zelf blijft gewoon bestaan.`}
          onConfirm={() => doeMarkeerOpen(openZonderKoppeling)} onCancel={() => setOpenZonderKoppeling(null)} />
      )}
    </div>
  );
}

function ContributieForm({ member, jaar, onSave, onClose }) {
  const [f, setF] = useState({ bedrag: '', datum: new Date().toISOString().slice(0, 10), rekening: '908', boeken: true });
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  return (
    <Modal title={`Contributie ${jaar} — ${fullName(member)}`} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Bedrag"><input type="number" step="0.01" className={inputCls} style={inputStyle} value={f.bedrag} onChange={e => upd('bedrag', e.target.value)} /></Field>
        <Field label="Datum ontvangen"><input type="date" className={inputCls} style={inputStyle} value={f.datum} onChange={e => upd('datum', e.target.value)} /></Field>
        <label className="flex items-start gap-2 text-sm rounded-lg p-2.5" style={{ background: C.paperDim }}>
          <input type="checkbox" checked={f.boeken} onChange={e => upd('boeken', e.target.checked)} className="mt-0.5" />
          <span>Ook automatisch boeken in Financiën</span>
        </label>
        {f.boeken && (
          <Field label="Op rekening">
            <select className={inputCls} style={inputStyle} value={f.rekening} onChange={e => upd('rekening', e.target.value)}>
              <option value="908">.908 lopend</option>
              <option value="319">.319 spaar</option>
            </select>
          </Field>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(f)} disabled={!f.bedrag}>Markeer betaald</Btn>
      </div>
    </Modal>
  );
}

function BulkWhatsAppModal({ leden, onClose }) {
  const [tekst, setTekst] = useState('');
  const [verzonden, setVerzonden] = useState([]);
  const zonderTelefoon = leden.filter(m => !m.telefoon);
  const metTelefoon = leden.filter(m => m.telefoon);

  function markeerVerzonden(id) {
    setVerzonden(v => v.includes(id) ? v : [...v, id]);
  }

  return (
    <Modal title="WhatsApp-bericht sturen" onClose={onClose} wide>
      <div className="space-y-3">
        <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>
          WhatsApp ondersteunt geen echt gelijktijdig verzenden naar meerdere losse nummers vanuit de browser. Schrijf hieronder één bericht en klik daarna per lid op "Verstuur" — dat opent WhatsApp met het bericht al klaargezet, jij hoeft alleen nog op versturen te tikken.
        </p>
        <Field label="Bericht">
          <textarea rows={3} className={inputCls} style={inputStyle} value={tekst} onChange={e => setTekst(e.target.value)} placeholder="Typ hier je bericht…" />
        </Field>
        <div>
          <span className="block text-xs font-medium mb-1.5" style={{ color: C.inkSoft }}>{metTelefoon.length} lid/leden met telefoonnummer</span>
          <div className="space-y-1.5 max-h-64 overflow-y-auto">
            {metTelefoon.map(m => (
              <div key={m.id} className="flex items-center justify-between gap-2 px-2.5 py-1.5 rounded-lg" style={{ background: C.paperDim }}>
                <div>
                  <span className="text-sm">{fullName(m)}</span>
                  <span className="text-xs ml-2" style={{ color: C.inkSoft }}>{m.telefoon}</span>
                </div>
                <a href={waLink(m.telefoon, tekst)} target="_blank" rel="noreferrer" onClick={() => markeerVerzonden(m.id)}>
                  <Badge tone={verzonden.includes(m.id) ? 'sage' : 'clay'}>{verzonden.includes(m.id) ? 'Verzonden ✓' : 'Verstuur'}</Badge>
                </a>
              </div>
            ))}
          </div>
        </div>
        {zonderTelefoon.length > 0 && (
          <p className="text-xs" style={{ color: C.rose }}>Geen telefoonnummer bekend van: {zonderTelefoon.map(m => fullName(m)).join(', ')}.</p>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Sluiten</Btn>
      </div>
    </Modal>
  );
}

function MemberForm({ member, dagdelen, onSave, onClose }) {
  const [f, setF] = useState(() => member ? { ...member } : {
    id: null, voornaam: '', tussenvoegsel: '', achternaam: '', email: '', telefoon: '',
    adres: '', postcode: '', woonplaats: '', gebdatum: '', lidsinds: '', eindelidmaat: '',
    status: 'actief', functie: '', dagdelen: [], groepsapp: 'ja',
  });
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  function toggleSlot(slot) {
    setF(s => {
      const has = (s.dagdelen || []).includes(slot);
      return { ...s, dagdelen: has ? s.dagdelen.filter(x => x !== slot) : [...(s.dagdelen || []), slot] };
    });
  }
  return (
    <Modal title={member ? 'Lid bewerken' : 'Nieuw lid'} onClose={onClose} wide>
      <div className="grid sm:grid-cols-2 gap-3">
        <Field label="Voornaam"><input className={inputCls} style={inputStyle} value={f.voornaam || ''} onChange={e => upd('voornaam', e.target.value)} /></Field>
        <Field label="Achternaam"><input className={inputCls} style={inputStyle} value={f.achternaam || ''} onChange={e => upd('achternaam', e.target.value)} /></Field>
        <Field label="E-mailadres"><input className={inputCls} style={inputStyle} value={f.email || ''} onChange={e => upd('email', e.target.value)} /></Field>
        <Field label="Telefoonnr."><input className={inputCls} style={inputStyle} value={f.telefoon || ''} onChange={e => upd('telefoon', e.target.value)} /></Field>
        <Field label="Adres"><input className={inputCls} style={inputStyle} value={f.adres || ''} onChange={e => upd('adres', e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-2">
          <Field label="Postcode"><input className={inputCls} style={inputStyle} value={f.postcode || ''} onChange={e => upd('postcode', e.target.value)} /></Field>
          <Field label="Woonplaats"><input className={inputCls} style={inputStyle} value={f.woonplaats || ''} onChange={e => upd('woonplaats', e.target.value)} /></Field>
        </div>
        <Field label="Lid sinds"><input type="date" className={inputCls} style={inputStyle} value={f.lidsinds || ''} onChange={e => upd('lidsinds', e.target.value)} /></Field>
        <Field label="Functie"><input className={inputCls} style={inputStyle} placeholder="bv. Bestuurslid" value={f.functie || ''} onChange={e => upd('functie', e.target.value)} /></Field>
        <Field label="Status">
          <select className={inputCls} style={inputStyle} value={f.status || 'actief'} onChange={e => upd('status', e.target.value)}>
            <option value="actief">Actief</option>
            <option value="inactief">Inactief</option>
          </select>
        </Field>
        <Field label="Groepsapp">
          <select className={inputCls} style={inputStyle} value={f.groepsapp || 'ja'} onChange={e => upd('groepsapp', e.target.value)}>
            {GROEPSAPP_OPTIES.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
          </select>
        </Field>
        {f.status === 'inactief' && (
          <Field label="Lidmaatschap beëindigd op"><input type="date" className={inputCls} style={inputStyle} value={f.eindelidmaat || ''} onChange={e => upd('eindelidmaat', e.target.value)} /></Field>
        )}
      </div>
      <div className="mt-3">
        <span className="block text-xs font-medium mb-1.5" style={{ color: C.inkSoft }}>Dagdelen / groepsindeling</span>
        <div className="flex flex-wrap gap-1.5">
          {dagdelen.map(slot => {
            const on = (f.dagdelen || []).includes(slot);
            return (
              <button key={slot} type="button" onClick={() => toggleSlot(slot)}
                className="px-2.5 py-1 rounded-full text-xs border"
                style={{ background: on ? C.clay : 'transparent', color: on ? '#fff' : C.ink, borderColor: on ? C.clay : C.border }}>
                {slot}
              </button>
            );
          })}
        </div>
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(f)}>Opslaan</Btn>
      </div>
    </Modal>
  );
}

/* ---------------------------- Ledenoverzichten --------------------------- */
const LEDEN_VELDEN = [
  { key: 'voornaam', label: 'Voornaam' },
  { key: 'achternaam', label: 'Achternaam' },
  { key: 'email', label: 'E-mailadres' },
  { key: 'telefoon', label: 'Telefoonnr.' },
  { key: 'adres', label: 'Adres' },
  { key: 'postcode', label: 'Postcode' },
  { key: 'woonplaats', label: 'Woonplaats' },
  { key: 'gebdatum', label: 'Geboortedatum' },
  { key: 'lidsinds', label: 'Lid sinds' },
  { key: 'eindelidmaat', label: 'Einde lidmaatschap' },
  { key: 'status', label: 'Status' },
  { key: 'functie', label: 'Functie' },
  { key: 'dagdelen', label: 'Dagdelen / groepen' },
  { key: 'groepsapp', label: 'Groepsapp' },
];
const DATE_VELDEN = new Set(['gebdatum', 'lidsinds', 'eindelidmaat']);
const GROEPSAPP_OPTIES = [
  { value: 'ja', label: 'Ja' },
  { value: 'afgemeld', label: 'Afgemeld' },
  { value: 'geen_mobiel', label: 'Geen mobiel' },
];

function ledenVeldWaarde(m, key) {
  if (key === 'dagdelen') return (m.dagdelen || []).join(', ') || '—';
  if (key === 'groepsapp') return (GROEPSAPP_OPTIES.find(o => o.value === m.groepsapp) || {}).label || 'Ja';
  if (DATE_VELDEN.has(key)) return m[key] ? fmtDate(m[key]) : '—';
  return m[key] || '—';
}

function ReportsModal({ members, dagdelen, onClose }) {
  const [soort, setSoort] = useState('leden');
  const [status, setStatus] = useState('actief');
  const [groepen, setGroepen] = useState([...dagdelen]);
  const [velden, setVelden] = useState(['voornaam', 'achternaam', 'email', 'telefoon', 'adres', 'status']);

  function toggleGroep(slot) {
    setGroepen(g => g.includes(slot) ? g.filter(x => x !== slot) : [...g, slot]);
  }
  function toggleVeld(key) {
    setVelden(v => v.includes(key) ? v.filter(x => x !== key) : [...v, key]);
  }

  const basis = members.filter(m => status === 'alle' || m.status === status);
  // Groepenfilter: bij Groepsamenstelling altijd toepassen (dat is het hele punt van dat overzicht).
  // Bij Ledenlijst alleen toepassen als de gebruiker de groepenselectie bewust heeft versmald —
  // anders zouden leden zonder (nog) toegewezen dagdeel (bv. inactieve leden) ten onrechte wegvallen.
  const groepenFilterActief = soort === 'groepen' || groepen.length < dagdelen.length;
  const geselecteerdeLeden = basis.filter(m => !groepenFilterActief || (m.dagdelen || []).some(d => groepen.includes(d)))
    .sort((a, b) => fullName(a).localeCompare(fullName(b)));

  function ledenlijstData() {
    const cols = LEDEN_VELDEN.filter(v => velden.includes(v.key));
    const header = cols.map(c => c.label);
    const rows = geselecteerdeLeden.map(m => cols.map(c => ledenVeldWaarde(m, c.key)));
    return { header, rows, titel: 'Ledenlijst' };
  }
  function groepsamenstellingData() {
    const cols = LEDEN_VELDEN.filter(v => velden.includes(v.key));
    const gekozenGroepen = dagdelen.filter(s => groepen.includes(s));
    const header = [...cols.map(c => c.label), ...gekozenGroepen];
    const rows = geselecteerdeLeden.map(m => [
      ...cols.map(c => ledenVeldWaarde(m, c.key)),
      ...gekozenGroepen.map(g => (m.dagdelen || []).includes(g) ? '✓' : ''),
    ]);
    return { header, rows, titel: 'Groepsamenstelling' };
  }

  function huidigeData() {
    return soort === 'leden' ? ledenlijstData() : groepsamenstellingData();
  }

  function exportExcel() {
    const { header, rows, titel } = huidigeData();
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([[titel], [`Groepen: ${groepen.length === dagdelen.length ? 'alle' : groepen.join(', ')}`], [], header, ...rows]);
    XLSX.utils.book_append_sheet(wb, ws, titel.slice(0, 31));
    XLSX.writeFile(wb, `BladelsCreatief_${titel.replace(/\s+/g, '_')}.xlsx`);
  }
  function exportPdf() {
    const { header, rows, titel } = huidigeData();
    const subtitel = `Groepen: ${groepen.length === dagdelen.length ? 'alle' : (groepen.join(', ') || 'geen')} · Status: ${status === 'alle' ? 'alle' : status} · ${rows.length} leden · ${new Date().toLocaleDateString('nl-NL')}`;
    downloadPrintableHtml({ titel, subtitel, header, rows });
    onClose();
  }

  return (
    <Modal title="Overzicht maken" onClose={onClose} wide>
      <div className="space-y-4">
        <div className="flex gap-1">
          {[['leden', 'Ledenlijst'], ['groepen', 'Groepsamenstelling']].map(([id, label]) => (
            <button key={id} onClick={() => setSoort(id)} className="px-3 py-1.5 rounded-lg text-sm font-medium border"
              style={{ background: soort === id ? C.clay : 'transparent', color: soort === id ? '#fff' : C.ink, borderColor: soort === id ? C.clay : C.border }}>
              {label}
            </button>
          ))}
        </div>

        <div>
          <span className="block text-xs font-medium mb-1.5" style={{ color: C.inkSoft }}>Status</span>
          <select value={status} onChange={e => setStatus(e.target.value)} className={inputCls} style={{ ...inputStyle, width: 180 }}>
            <option value="actief">Alleen actieve leden</option>
            <option value="inactief">Alleen inactieve leden</option>
            <option value="alle">Alle leden</option>
          </select>
        </div>

        <div>
          <div className="flex items-center justify-between mb-1.5">
            <span className="text-xs font-medium" style={{ color: C.inkSoft }}>Voor welke groepen moet het overzicht gemaakt worden?</span>
            <div className="flex gap-2 text-xs">
              <button onClick={() => setGroepen([...dagdelen])} className="underline" style={{ color: C.clay }}>alles</button>
              <button onClick={() => setGroepen([])} className="underline" style={{ color: C.inkSoft }}>niets</button>
            </div>
          </div>
          <div className="flex flex-wrap gap-1.5">
            {dagdelen.map(slot => {
              const on = groepen.includes(slot);
              return (
                <button key={slot} type="button" onClick={() => toggleGroep(slot)}
                  className="px-2.5 py-1 rounded-full text-xs border"
                  style={{ background: on ? C.clay : 'transparent', color: on ? '#fff' : C.ink, borderColor: on ? C.clay : C.border }}>
                  {slot}
                </button>
              );
            })}
          </div>
        </div>

        <div>
          <span className="block text-xs font-medium mb-1.5" style={{ color: C.inkSoft }}>Welke velden moeten in het overzicht staan?</span>
          <div className="grid sm:grid-cols-3 gap-1.5">
            {LEDEN_VELDEN.map(v => (
              <label key={v.key} className="flex items-center gap-1.5 text-sm">
                <input type="checkbox" checked={velden.includes(v.key)} onChange={() => toggleVeld(v.key)} />
                {v.label}
              </label>
            ))}
          </div>
          {soort === 'groepen' && (
            <p className="text-xs mt-1.5" style={{ color: C.inkSoft }}>
              Deze velden vormen de eerste kolommen; daarna volgt een kolom per hierboven geselecteerde groep met een vinkje bij deelname.
            </p>
          )}
        </div>

        <p className="text-xs" style={{ color: C.inkSoft }}>{geselecteerdeLeden.length} leden komen in dit overzicht.</p>

        <div className="rounded-lg px-3 py-2 flex gap-2 items-start" style={{ background: C.paperDim }}>
          <Printer size={14} style={{ color: C.inkSoft, marginTop: 2, flexShrink: 0 }} />
          <p className="text-xs" style={{ color: C.inkSoft }}>
            <strong>Als PDF</strong> probeert een opmaakklaar overzicht in een nieuw tabblad te openen, waar automatisch het printvenster verschijnt — kies daar bij <em>Bestemming/Printer</em> de optie <strong>"Opslaan als PDF"</strong>. Blokkeert je browser pop-ups, dan wordt in plaats daarvan een bestand gedownload dat je zelf even opent (bijv. vanuit je Downloads-map) — ook dan opent het printvenster automatisch.
            <br /><strong>Als Excel</strong> downloadt direct een .xlsx-bestand, zonder tussenstap.
          </p>
        </div>

        <div className="flex justify-end gap-2 pt-2 border-t" style={{ borderColor: C.border }}>
          <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
          <Btn tone="outline" icon={Printer} onClick={exportPdf} disabled={!geselecteerdeLeden.length} title="Downloadt een printklaar bestand met automatisch printvenster">Als PDF</Btn>
          <Btn tone="sage" icon={Download} onClick={exportExcel} disabled={!geselecteerdeLeden.length}>Als Excel</Btn>
        </div>
      </div>
    </Modal>
  );
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* Bouwt een Word-compatibel .doc bestand (HTML met Word-namespaces) — opent direct in Word,
   waar het meteen als volwaardig document bewerkt en desgewenst als .docx opgeslagen kan worden. */
/* Bouwt de <img>-tag voor het logo in Word-exports, met de ingestelde hoogte (cm). Naast de
   CSS-hoogte ook expliciete pixel-attributen, want Word negeert bij MHTML-ingesloten
   afbeeldingen anders de CSS-hoogte en gebruikt het de eigen pixelformaat van de afbeelding. */
function logoImgTag(hoogteCm) {
  const h = Math.round((hoogteCm / 2.54) * 96);
  const w = Math.round(h * (400 / 208));
  return `<img src="${LOGO_URI}" width="${w}" height="${h}" style="height:${hoogteCm}cm;max-height:${hoogteCm}cm;width:${w}px;display:block;margin-bottom:4px;" />`;
}
function downloadWordDoc({ titel, filename, bodyHtml }) {
  // Word rendert ingesloten data-URI-afbeeldingen (base64 rechtstreeks in <img src>) onbetrouwbaar.
  // We vervangen het logo daarom door een cid-verwijzing en bouwen een MHTML-bestand (multipart),
  // het formaat dat Word wél altijd correct gebruikt voor ingesloten afbeeldingen.
  const logoBase64 = LOGO_URI.split(',')[1] || '';
  const bodyMetCid = bodyHtml.split(LOGO_URI).join('cid:logo.png');
  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(titel)}</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
  @page { margin: 2.2cm; }
  body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #1a1410; }
  h1 { font-size: 19pt; color: ${C.clayDeep}; margin: 10px 0 2px; }
  h2 { font-size: 13.5pt; color: ${C.clay}; border-bottom: 1px solid #cccccc; padding-bottom: 3px; margin-top: 26px; }
  h3 { font-size: 11.5pt; margin: 14px 0 3px; color: #1a1410; }
  p { margin: 4px 0; line-height: 1.4; }
  table { border-collapse: collapse; width: 100%; margin: 6px 0 12px; }
  th, td { border: 1px solid #aaaaaa; padding: 5px 8px; font-size: 10.5pt; text-align: left; vertical-align: top; }
  th { background: #EEF3F9; }
  .meta { color: #555555; font-size: 10pt; margin-bottom: 10px; }
  .status-open { color: ${C.rose}; font-weight: bold; }
  .status-klaar { color: ${C.sageDeep}; font-weight: bold; }
  .bar { height: 3px; background: linear-gradient(90deg, ${C.clayDeep}, ${C.clay}, ${C.sage}, ${C.ochre}, ${C.rose}); margin: 4px 0 16px; }
</style>
</head>
<body>
${bodyMetCid}
</body>
</html>`;

  const boundary = "----BladelsCreatiefMHTBoundary";
  const mhtml = [
    'MIME-Version: 1.0',
    `Content-Type: multipart/related; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    'Content-Type: text/html; charset="utf-8"',
    'Content-Location: file:///document.html',
    '',
    html,
    '',
    `--${boundary}`,
    'Content-Type: image/png',
    'Content-Transfer-Encoding: base64',
    'Content-Location: logo.png',
    'Content-ID: <logo.png>',
    '',
    logoBase64,
    '',
    `--${boundary}--`,
    '',
  ].join('\r\n');

  const blob = new Blob([mhtml], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

function downloadPrintableHtml({ titel, subtitel, header, rows }) {
  const theadHtml = header.map(h => `<th>${escapeHtml(h)}</th>`).join('');
  const rowsHtml = rows.map((r, i) => `<tr style="background:${i % 2 ? '#F7F2E6' : 'transparent'}">${r.map(c => `<td>${escapeHtml(c)}</td>`).join('')}</tr>`).join('');
  const html = `<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(titel)} — BladelsCreatief</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: Arial, Helvetica, sans-serif; color: #1a1410; padding: 32px; max-width: 900px; margin: 0 auto; }
  img.logo { height: 46px; display: block; margin-bottom: 6px; }
  .bar { height: 3px; width: 100%; background: linear-gradient(90deg, ${C.clayDeep}, ${C.clay}, ${C.sage}, ${C.ochre}, ${C.rose}); margin-bottom: 18px; border-radius: 2px; }
  h1 { font-size: 21px; margin: 0 0 2px; }
  p.sub { font-size: 12px; color: #5A6B8C; margin: 0 0 16px; }
  p.foot { font-size: 10px; color: #5A6B8C; margin-top: 18px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th { text-align: left; border-bottom: 1.5px solid #13244A; padding: 6px 9px; font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.3px; }
  td { padding: 6px 9px; border-bottom: 1px solid #DCE6F2; }
  .tip { font-size: 12px; background: #E9F1F9; border-radius: 8px; padding: 10px 14px; margin-bottom: 18px; color: #5A6B8C; }
  @media print { .tip { display: none; } body { padding: 0; max-width: none; } }
</style>
</head>
<body>
  <div class="tip">Tip: gebruik <strong>Ctrl/Cmd + P</strong> en kies "Opslaan als PDF" om dit overzicht als pdf te bewaren. Dit printvenster opent hieronder ook automatisch.</div>
  <img class="logo" src="${LOGO_URI}" alt="BladelsCreatief" />
  <div class="bar"></div>
  <h1>${escapeHtml(titel)}</h1>
  ${subtitel ? `<p class="sub">${escapeHtml(subtitel)}</p>` : ''}
  <table>
    <thead><tr>${theadHtml}</tr></thead>
    <tbody>${rowsHtml}</tbody>
  </table>
  <p class="foot">Gegenereerd op ${new Date().toLocaleDateString('nl-NL')} · ${rows.length} rijen</p>
  <script>window.addEventListener('load', function () { setTimeout(function () { try { window.print(); } catch (e) {} }, 200); });</script>
</body>
</html>`;

  const blob = new Blob([html], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  const win = window.open(url, '_blank');
  if (!win) {
    // Pop-up geblokkeerd: val terug op een downloadbaar bestand dat de gebruiker zelf opent.
    const a = document.createElement('a');
    a.href = url;
    a.download = `BladelsCreatief_${titel.replace(/\s+/g, '_')}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
  setTimeout(() => URL.revokeObjectURL(url), 8000);
}

/* Zelfde print-opzet als downloadPrintableHtml, maar met vrije (meerdere secties/tabellen)
   inhoud in plaats van één platte tabel — voor overzichten zoals de workshops-samenvatting. */
/* Bouwt de kerninhoud van het workshopoverzicht (per workshop: gegevens + deelnemerslijst) —
   zonder logo, want de PDF-wrapper voegt die zelf al toe; de Word-export zet 'm er apart voor. */
function workshopsOverzichtKern(workshops, inschrijvingen) {
  const secties = workshops.map(w => {
    const deelnemers = inschrijvingen.filter(i => i.workshopId === w.id)
      .sort((a, b) => (a.status === b.status ? 0 : a.status === 'wachtlijst' ? 1 : -1) || (a.naam || '').localeCompare(b.naam || ''));
    const rijen = deelnemers.map(i => `<tr><td>${escapeHtml(i.naam || '')}</td><td>${escapeHtml([i.email, i.telefoon].filter(Boolean).join(' · ') || '—')}</td><td>${i.status === 'wachtlijst' ? 'Wachtlijst' : 'Ingeschreven'}</td><td>${i.betaald ? 'Betaald' : 'Open'}</td></tr>`).join('');
    return `<h2>${escapeHtml(w.titel)}</h2>
<p class="meta">${escapeHtml(w.soort)} · ${w.type === 'reeks' ? `Reeks (${(w.datums || []).length} data)` : 'Eenmalig'}${(w.datums || []).length ? ' · ' + w.datums.map(fmtDate).join(', ') : ''}${w.dagdeel ? ' · ' + escapeHtml(w.dagdeel) : ''}${w.locatie ? ' · ' + escapeHtml(w.locatie) : ''}</p>
<p class="meta">${w.bedrag != null ? euro(w.bedrag) + ' per deelnemer' : 'Geen bedrag ingesteld'}${w.maxDeelnemers != null ? ` · max. ${w.maxDeelnemers} deelnemers` : ''} · status: ${escapeHtml(w.status)} · ${deelnemers.length} deelnemer(s)</p>
<table>
  <thead><tr><th>Naam</th><th>Contact</th><th>Status</th><th>Betaald</th></tr></thead>
  <tbody>${rijen || '<tr><td colspan="4"><em>Geen deelnemers ingeschreven.</em></td></tr>'}</tbody>
</table>`;
  }).join('');
  return `<h1>Workshopoverzicht</h1>
<p class="meta">Gegenereerd op ${new Date().toLocaleDateString('nl-NL')} · ${workshops.length} workshop(s)</p>
${secties}`;
}

function downloadPrintableHtmlVrij({ titel, bodyHtml }) {
  const html = `<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(titel)} — BladelsCreatief</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: Arial, Helvetica, sans-serif; color: #1a1410; padding: 32px; max-width: 900px; margin: 0 auto; }
  img.logo { height: 46px; display: block; margin-bottom: 6px; }
  .bar { height: 3px; width: 100%; background: linear-gradient(90deg, ${C.clayDeep}, ${C.clay}, ${C.sage}, ${C.ochre}, ${C.rose}); margin-bottom: 18px; border-radius: 2px; }
  h1 { font-size: 21px; margin: 0 0 2px; }
  h2 { font-size: 15px; margin: 22px 0 4px; border-bottom: 1.5px solid #DCE6F2; padding-bottom: 4px; }
  p.sub, p.meta { font-size: 12px; color: #5A6B8C; margin: 0 0 10px; }
  p.foot { font-size: 10px; color: #5A6B8C; margin-top: 18px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 10px; }
  th { text-align: left; border-bottom: 1.5px solid #13244A; padding: 6px 9px; font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.3px; }
  td { padding: 6px 9px; border-bottom: 1px solid #DCE6F2; }
  .tip { font-size: 12px; background: #E9F1F9; border-radius: 8px; padding: 10px 14px; margin-bottom: 18px; color: #5A6B8C; }
  @media print { .tip { display: none; } body { padding: 0; max-width: none; } h2 { page-break-before: auto; } }
</style>
</head>
<body>
  <div class="tip">Tip: gebruik <strong>Ctrl/Cmd + P</strong> en kies "Opslaan als PDF" om dit overzicht als pdf te bewaren. Dit printvenster opent hieronder ook automatisch.</div>
  <img class="logo" src="${LOGO_URI}" alt="BladelsCreatief" />
  <div class="bar"></div>
  ${bodyHtml}
  <script>window.addEventListener('load', function () { setTimeout(function () { try { window.print(); } catch (e) {} }, 200); });</script>
</body>
</html>`;

  const blob = new Blob([html], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  const win = window.open(url, '_blank');
  if (!win) {
    const a = document.createElement('a');
    a.href = url;
    a.download = `BladelsCreatief_${titel.replace(/\s+/g, '_')}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
  setTimeout(() => URL.revokeObjectURL(url), 8000);
}

/* =========================================================================
   JEUGDATELIER
========================================================================= */
/* =========================================================================
   VERGADERINGEN — agenda, notulen & actielijst gekoppeld, export naar Word
========================================================================= */
function vergaderingWordHtml(v, actielijst, logoHoogteCm) {
  const gekoppeld = actielijst.filter(a => a.vergaderingId === v.id);
  const agendaLijst = v.agendapunten.map(p => `<li>${escapeHtml(p.titel)}${p.toelichting ? ' — <em>' + escapeHtml(p.toelichting) + '</em>' : ''}</li>`).join('');
  const notulenHtml = v.agendapunten.map((p, i) => {
    const acties = actielijst.filter(a => a.agendapuntId === p.id);
    const actiesHtml = acties.length ? `<table><tr><th>Actiepunt</th><th>Wie</th><th>Deadline</th><th>Status</th></tr>${acties.map(a => `<tr><td>${escapeHtml(a.omschrijving)}</td><td>${escapeHtml(a.wie || '—')}</td><td>${a.deadline ? fmtDate(a.deadline) : '—'}</td><td class="${a.status === 'klaar' ? 'status-klaar' : 'status-open'}">${a.status === 'klaar' ? 'Klaar' : 'Open'}</td></tr>`).join('')}</table>` : '';
    return `<h3>${i + 1}. ${escapeHtml(p.titel)}</h3><p>${escapeHtml(p.notulen || '(geen notulen ingevoerd)').replace(/\n/g, '<br/>')}</p>${actiesHtml}`;
  }).join('');
  const actieTabel = gekoppeld.length
    ? `<table><tr><th>Actiepunt</th><th>Agendapunt</th><th>Wie</th><th>Deadline</th><th>Status</th></tr>${gekoppeld.map(a => `<tr><td>${escapeHtml(a.omschrijving)}</td><td>${escapeHtml(a.agendapuntTitel || '—')}</td><td>${escapeHtml(a.wie || '—')}</td><td>${a.deadline ? fmtDate(a.deadline) : '—'}</td><td class="${a.status === 'klaar' ? 'status-klaar' : 'status-open'}">${a.status === 'klaar' ? 'Klaar' : 'Open'}</td></tr>`).join('')}</table>`
    : '<p><em>Geen actiepunten vastgelegd bij deze vergadering.</em></p>';

  return `${logoImgTag(logoHoogteCm)}<div class="bar"></div>
<h1>${escapeHtml(v.titel)}</h1>
<p class="meta">${fmtDate(v.datum)}${v.locatie ? ' · ' + escapeHtml(v.locatie) : ''}${v.aanwezigen ? '<br/>Aanwezig: ' + escapeHtml(v.aanwezigen) : ''}</p>
<h2>Agenda</h2>
<ol>${agendaLijst || '<li><em>Geen agendapunten</em></li>'}</ol>
<h2>Notulen</h2>
${notulenHtml || '<p><em>Nog geen notulen ingevoerd.</em></p>'}
<h2>Actiepunten — koppeling met agenda &amp; notulen</h2>
<p>Onderstaande actiepunten zijn direct gekoppeld aan het agendapunt waar ze tijdens deze vergadering uit voortkwamen.</p>
${actieTabel}`;
}

function actielijstWordHtml(items, logoHoogteCm) {
  const rows = items.map(a => `<tr><td>${escapeHtml(a.omschrijving)}</td><td>${escapeHtml(a.vergaderingTitel || '—')}</td><td>${escapeHtml(a.agendapuntTitel || '—')}</td><td>${escapeHtml(a.wie || '—')}</td><td>${a.deadline ? fmtDate(a.deadline) : '—'}</td><td class="${a.status === 'klaar' ? 'status-klaar' : 'status-open'}">${a.status === 'klaar' ? 'Klaar' : 'Open'}</td></tr>`).join('');
  return `${logoImgTag(logoHoogteCm)}<div class="bar"></div>
<h1>Actielijst</h1>
<p class="meta">Gegenereerd op ${new Date().toLocaleDateString('nl-NL')} · ${items.length} actiepunten</p>
<table><tr><th>Actiepunt</th><th>Vergadering</th><th>Agendapunt</th><th>Wie</th><th>Deadline</th><th>Status</th></tr>${rows || '<tr><td colspan="6"><em>Geen actiepunten</em></td></tr>'}</table>`;
}

function VergaderingenTab({ members, vergaderingen, setVergaderingen, actielijst, setActielijst, agendapuntenVooraf, agendapuntenAfsluitend, logoHoogteCm, readOnly, onTrash, onLog }) {
  const [view, setView] = useState('lijst');
  const [selectedId, setSelectedId] = useState(null);
  const selected = vergaderingen.find(v => v.id === selectedId);

  if (view === 'detail' && selected) {
    return <MeetingDetail members={members} vergadering={selected} vergaderingen={vergaderingen} setVergaderingen={setVergaderingen}
      actielijst={actielijst} setActielijst={setActielijst} logoHoogteCm={logoHoogteCm} onBack={() => setView('lijst')} readOnly={readOnly} onLog={onLog} />;
  }
  if (view === 'actielijst') {
    return <ActielijstView actielijst={actielijst} setActielijst={setActielijst} logoHoogteCm={logoHoogteCm} onBack={() => setView('lijst')} readOnly={readOnly} />;
  }
  return (
    <MeetingList members={members} vergaderingen={vergaderingen} setVergaderingen={setVergaderingen} actielijst={actielijst} setActielijst={setActielijst}
      agendapuntenVooraf={agendapuntenVooraf} agendapuntenAfsluitend={agendapuntenAfsluitend}
      onOpen={id => { setSelectedId(id); setView('detail'); }} onOpenActielijst={() => setView('actielijst')} readOnly={readOnly} onTrash={onTrash} onLog={onLog} />
  );
}

function AanwezigenField({ value, onChange, kandidaten }) {
  const lijst = (value || '').split(',').map(s => s.trim()).filter(Boolean);
  function toggle(naam) {
    const has = lijst.includes(naam);
    const nieuw = has ? lijst.filter(n => n !== naam) : [...lijst, naam];
    onChange(nieuw.join(', '));
  }
  return (
    <div>
      {kandidaten.length > 0 && (
        <div className="mb-2">
          <span className="block text-xs font-medium mb-1" style={{ color: C.inkSoft }}>Bestuursleden / leden met functie</span>
          <div className="flex flex-wrap gap-1.5">
            {kandidaten.map(m => {
              const naam = fullName(m);
              const on = lijst.includes(naam);
              return (
                <button key={m.id} type="button" onClick={() => toggle(naam)}
                  className="px-2.5 py-1 rounded-full text-xs border flex items-center gap-1"
                  style={{ background: on ? C.clay : 'transparent', color: on ? '#fff' : C.ink, borderColor: on ? C.clay : C.border }}>
                  {on ? <CheckSquare size={12} /> : <Square size={12} />} {naam}{m.functie ? ` (${m.functie})` : ''}
                </button>
              );
            })}
          </div>
        </div>
      )}
      <Field label="Aanwezigen (namen, gescheiden door komma's — vul hier ook gasten of overigen aan)">
        <DebouncedField value={value || ''} onCommit={onChange} />
      </Field>
    </div>
  );
}

function MeetingList({ members, vergaderingen, setVergaderingen, actielijst, setActielijst, agendapuntenVooraf, agendapuntenAfsluitend, onOpen, onOpenActielijst, readOnly, onTrash, onLog }) {
  const [showNew, setShowNew] = useState(false);
  const [delId, setDelId] = useState(null);
  const openActies = actielijst.filter(a => a.status !== 'klaar');
  const sorted = [...vergaderingen].sort((a, b) => (b.datum || '').localeCompare(a.datum || ''));

  function createMeeting(f, neemOver, standaardPunten) {
    const id = uid(vergaderingen);
    let volgId = 1;
    const agendapunten = [];
    function voegToe(titel, toelichting) {
      const puntId = volgId++;
      agendapunten.push({ id: puntId, titel, toelichting: toelichting || '', notulen: '' });
      return puntId;
    }
    if (standaardPunten) agendapuntenVooraf.forEach(t => voegToe(t));
    let actiepuntenPuntId = null;
    const actiepuntenTitel = 'Actiepunten vorige vergadering(en)';
    if (neemOver && openActies.length) {
      actiepuntenPuntId = voegToe(actiepuntenTitel, 'Zie de actiepunten hieronder — aan te vinken zodra afgerond.');
    }
    if (standaardPunten) agendapuntenAfsluitend.forEach(t => voegToe(t));
    setVergaderingen([...vergaderingen, { id, ...f, agendapunten }]);
    if (actiepuntenPuntId != null) {
      const openIds = new Set(openActies.map(a => a.id));
      setActielijst(actielijst.map(a => openIds.has(a.id)
        ? { ...a, vergaderingId: id, vergaderingTitel: f.titel, agendapuntId: actiepuntenPuntId, agendapuntTitel: actiepuntenTitel }
        : a));
    }
    onLog(`Vergadering aangemaakt: ${f.titel} (${f.datum})`, 'vergaderingen');
    setShowNew(false);
    onOpen(id);
  }
  function removeMeeting(id) {
    const v = vergaderingen.find(x => x.id === id);
    setVergaderingen(vergaderingen.filter(v => v.id !== id));
    if (v) { onTrash('vergadering', v); onLog(`Vergadering verwijderd: ${v.titel} (${v.datum})`, 'vergaderingen'); }
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <h2 className="font-semibold" style={{ fontFamily: 'Fraunces, serif', fontSize: 18, color: C.ink }}>Vergaderingen</h2>
          <p className="text-xs" style={{ color: C.inkSoft }}>Agenda, notulen en actiepunten in samenhang — met export naar Word.</p>
        </div>
        <div className="flex gap-2">
          <Btn tone="outline" icon={CheckSquare} onClick={onOpenActielijst}>
            Actielijst{openActies.length ? ` (${openActies.length} open)` : ''}
          </Btn>
          {!readOnly && <Btn icon={Plus} onClick={() => setShowNew(true)}>Nieuwe vergadering</Btn>}
        </div>
      </div>

      <div className="grid sm:grid-cols-2 gap-3">
        {sorted.map(v => {
          const gekoppeld = actielijst.filter(a => a.vergaderingId === v.id);
          const open = gekoppeld.filter(a => a.status !== 'klaar').length;
          return (
            <Card key={v.id} className="p-4 cursor-pointer hover:shadow-md transition-shadow" onClick={() => onOpen(v.id)}>
              <div className="flex items-start justify-between">
                <div>
                  <h3 className="font-semibold" style={{ color: C.ink }}>{v.titel}</h3>
                  <p className="text-xs" style={{ color: C.inkSoft }}>{fmtDate(v.datum)}{v.locatie ? ' · ' + v.locatie : ''}</p>
                </div>
                <button onClick={e => { e.stopPropagation(); setDelId(v.id); }} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.rose, visibility: readOnly ? 'hidden' : 'visible' }}><Trash2 size={14} /></button>
              </div>
              <div className="flex gap-1.5 mt-2.5 flex-wrap">
                <Badge tone="muted">{v.agendapunten.length} agendapunten</Badge>
                {gekoppeld.length > 0 && <Badge tone={open ? 'clay' : 'sage'}>{gekoppeld.length} actiepunten{open ? ` · ${open} open` : ' · alles klaar'}</Badge>}
              </div>
            </Card>
          );
        })}
        {!sorted.length && (
          <div className="sm:col-span-2">
            <EmptyState icon={ClipboardList} text="Nog geen vergaderingen vastgelegd. Maak een nieuwe vergadering aan om agenda en notulen te koppelen." />
          </div>
        )}
      </div>

      {showNew && <NewMeetingModal members={members} openActies={openActies} agendapuntenVooraf={agendapuntenVooraf} agendapuntenAfsluitend={agendapuntenAfsluitend} onSave={createMeeting} onClose={() => setShowNew(false)} />}
      {delId != null && (
        <ConfirmModal message="Deze vergadering (met agenda en notulen) verwijderen? Gekoppelde actiepunten blijven bestaan in de actielijst." onConfirm={() => { removeMeeting(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
    </div>
  );
}

function NewMeetingModal({ members, openActies, agendapuntenVooraf, agendapuntenAfsluitend, onSave, onClose }) {
  const [f, setF] = useState({ datum: new Date().toISOString().slice(0, 10), titel: 'Bestuursvergadering', locatie: '', aanwezigen: '' });
  const [neemOver, setNeemOver] = useState(openActies.length > 0);
  const [standaardPunten, setStandaardPunten] = useState(true);
  const kandidaten = members.filter(m => m.functie && m.functie.trim()).sort((a, b) => fullName(a).localeCompare(fullName(b)));
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  return (
    <Modal title="Nieuwe vergadering" onClose={onClose} wide>
      <div className="space-y-3">
        <Field label="Titel"><input className={inputCls} style={inputStyle} value={f.titel} onChange={e => upd('titel', e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Datum"><input type="date" className={inputCls} style={inputStyle} value={f.datum} onChange={e => upd('datum', e.target.value)} /></Field>
          <Field label="Locatie"><input className={inputCls} style={inputStyle} value={f.locatie} onChange={e => upd('locatie', e.target.value)} /></Field>
        </div>
        <AanwezigenField value={f.aanwezigen} onChange={v => upd('aanwezigen', v)} kandidaten={kandidaten} />
        <label className="flex items-start gap-2 text-sm rounded-lg p-2.5" style={{ background: C.paperDim }}>
          <input type="checkbox" checked={standaardPunten} onChange={e => setStandaardPunten(e.target.checked)} className="mt-0.5" />
          <span>Standaard agendapunten toevoegen ({agendapuntenVooraf.join(', ')} … {agendapuntenAfsluitend.join(', ')})</span>
        </label>
        {openActies.length > 0 && (
          <label className="flex items-start gap-2 text-sm rounded-lg p-2.5" style={{ background: C.paperDim }}>
            <input type="checkbox" checked={neemOver} onChange={e => setNeemOver(e.target.checked)} className="mt-0.5" />
            <span>Neem de {openActies.length} openstaande actiepunten van vorige vergaderingen over als agendapunt</span>
          </label>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(f, neemOver, standaardPunten)} disabled={!f.titel || !f.datum}>Aanmaken</Btn>
      </div>
    </Modal>
  );
}

function MeetingDetail({ members, vergadering, vergaderingen, setVergaderingen, actielijst, setActielijst, logoHoogteCm, onBack, readOnly, onLog }) {
  const [actieVoorPunt, setActieVoorPunt] = useState(null);
  const [showAddPunt, setShowAddPunt] = useState(false);
  const [delPuntId, setDelPuntId] = useState(null);
  const [delActieId, setDelActieId] = useState(null);

  function update(patch) {
    setVergaderingen(vergaderingen.map(x => x.id === vergadering.id ? { ...x, ...patch } : x));
  }
  function updatePunt(puntId, patch) {
    update({ agendapunten: vergadering.agendapunten.map(p => p.id === puntId ? { ...p, ...patch } : p) });
  }
  function addPunt(titel) {
    update({ agendapunten: [...vergadering.agendapunten, { id: uid(vergadering.agendapunten), titel, toelichting: '', notulen: '' }] });
    onLog(`Agendapunt toegevoegd: ${titel} (${vergadering.titel})`, 'vergaderingen');
    setShowAddPunt(false);
  }
  function removePunt(puntId) {
    const p = vergadering.agendapunten.find(x => x.id === puntId);
    update({ agendapunten: vergadering.agendapunten.filter(p => p.id !== puntId) });
    if (p) onLog(`Agendapunt verwijderd: ${p.titel} (${vergadering.titel})`, 'vergaderingen');
  }
  function movePunt(puntId, richting) {
    const lijst = [...vergadering.agendapunten];
    const idx = lijst.findIndex(p => p.id === puntId);
    const nieuweIdx = idx + richting;
    if (idx < 0 || nieuweIdx < 0 || nieuweIdx >= lijst.length) return;
    [lijst[idx], lijst[nieuweIdx]] = [lijst[nieuweIdx], lijst[idx]];
    update({ agendapunten: lijst });
  }
  function addActie(punt, data) {
    const nieuw = {
      id: uid(actielijst), omschrijving: data.omschrijving, wie: data.wie, deadline: data.deadline, status: 'open',
      vergaderingId: vergadering.id, vergaderingTitel: vergadering.titel, agendapuntId: punt.id, agendapuntTitel: punt.titel,
    };
    setActielijst([...actielijst, nieuw]);
    onLog(`Actiepunt toegevoegd: ${data.omschrijving} (${vergadering.titel})`, 'vergaderingen');
    setActieVoorPunt(null);
  }
  function toggleActie(id) {
    setActielijst(actielijst.map(a => a.id === id ? { ...a, status: a.status === 'klaar' ? 'open' : 'klaar' } : a));
  }
  function removeActie(id) {
    setActielijst(actielijst.filter(a => a.id !== id));
  }
  function exportWord() {
    const html = vergaderingWordHtml(vergadering, actielijst, logoHoogteCm);
    downloadWordDoc({ titel: vergadering.titel, filename: `BladelsCreatief_${vergadering.titel.replace(/\s+/g, '_')}_${vergadering.datum}.doc`, bodyHtml: html });
  }

  return (
    <div className="space-y-4">
      <button onClick={onBack} className="flex items-center gap-1 text-sm" style={{ color: C.inkSoft }}><ArrowLeft size={15} /> Terug naar vergaderingen</button>

      <Card className="p-4">
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label="Titel"><DebouncedField disabled={readOnly} value={vergadering.titel} onCommit={v => update({ titel: v })} /></Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Datum"><input disabled={readOnly} type="date" className={inputCls} style={inputStyle} value={vergadering.datum} onChange={e => update({ datum: e.target.value })} /></Field>
            <Field label="Locatie"><DebouncedField disabled={readOnly} value={vergadering.locatie || ''} onCommit={v => update({ locatie: v })} /></Field>
          </div>
          <div className="sm:col-span-2">
            {readOnly ? (
              <Field label="Aanwezigen"><input disabled className={inputCls} style={inputStyle} value={vergadering.aanwezigen || ''} /></Field>
            ) : (
              <AanwezigenField value={vergadering.aanwezigen} onChange={v => update({ aanwezigen: v })}
                kandidaten={members.filter(m => m.functie && m.functie.trim()).sort((a, b) => fullName(a).localeCompare(fullName(b)))} />
            )}
          </div>
        </div>
        <div className="flex justify-end mt-3">
          <Btn tone="sage" icon={FileText} onClick={exportWord}>Als Word</Btn>
        </div>
      </Card>

      <div className="flex items-center justify-between">
        <h3 className="font-semibold flex items-center gap-1.5" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}><Link2 size={16} /> Agenda &amp; notulen</h3>
        {!readOnly && <Btn tone="outline" size="sm" icon={Plus} onClick={() => setShowAddPunt(true)}>Agendapunt toevoegen</Btn>}
      </div>

      <div className="space-y-3">
        {vergadering.agendapunten.map((p, i) => {
          const acties = actielijst.filter(a => a.agendapuntId === p.id);
          return (
            <Card key={p.id} className="p-4">
              <div className="flex items-start justify-between gap-2 mb-2">
                <div className="flex items-start gap-2 flex-1">
                  <Badge tone="clay">{i + 1}</Badge>
                  <DebouncedField disabled={readOnly} value={p.titel} onCommit={v => updatePunt(p.id, { titel: v })}
                    className="font-medium text-sm flex-1 rounded border px-2 py-1" style={inputStyle} />
                </div>
                {!readOnly && (
                  <div className="flex flex-col">
                    <button onClick={() => movePunt(p.id, -1)} disabled={i === 0} className="p-0.5 rounded hover:bg-black/5 disabled:opacity-20" style={{ color: C.inkSoft }}><ChevronUp size={14} /></button>
                    <button onClick={() => movePunt(p.id, 1)} disabled={i === vergadering.agendapunten.length - 1} className="p-0.5 rounded hover:bg-black/5 disabled:opacity-20" style={{ color: C.inkSoft }}><ChevronDown size={14} /></button>
                  </div>
                )}
                {!readOnly && <button onClick={() => setDelPuntId(p.id)} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={14} /></button>}
              </div>
              <Field label="Toelichting (optioneel, vooraf)">
                <DebouncedField disabled={readOnly} value={p.toelichting || ''} onCommit={v => updatePunt(p.id, { toelichting: v })} />
              </Field>
              <div className="mt-2">
                <Field label="Notulen">
                  <DebouncedField textarea disabled={readOnly} value={p.notulen || ''} onCommit={v => updatePunt(p.id, { notulen: v })} />
                </Field>
              </div>
              <div className="mt-2.5">
                <div className="flex items-center justify-between mb-1">
                  <span className="text-xs font-medium" style={{ color: C.inkSoft }}>Actiepunten bij dit agendapunt</span>
                  {!readOnly && <button onClick={() => setActieVoorPunt(p)} className="text-xs underline" style={{ color: C.clay }}>+ actiepunt</button>}
                </div>
                {acties.length > 0 ? (
                  <div className="space-y-1">
                    {acties.map(a => (
                      <div key={a.id} className="flex items-center gap-2 text-sm px-2 py-1.5 rounded" style={{ background: C.paperDim }}>
                        <button onClick={() => !readOnly && toggleActie(a.id)} disabled={readOnly} style={{ color: a.status === 'klaar' ? C.sageDeep : C.inkSoft }}>
                          {a.status === 'klaar' ? <CheckSquare size={15} /> : <Square size={15} />}
                        </button>
                        <span className="flex-1" style={{ textDecoration: a.status === 'klaar' ? 'line-through' : 'none', color: a.status === 'klaar' ? C.inkSoft : C.ink }}>{a.omschrijving}</span>
                        <span className="text-xs" style={{ color: C.inkSoft }}>{a.wie || '—'}</span>
                        <span className="text-xs" style={{ color: C.inkSoft }}>{a.deadline ? fmtDate(a.deadline) : ''}</span>
                        {!readOnly && <button onClick={() => setDelActieId(a.id)} style={{ color: C.rose }}><Trash2 size={12} /></button>}
                      </div>
                    ))}
                  </div>
                ) : <p className="text-xs italic" style={{ color: C.inkSoft }}>Geen actiepunten bij dit agendapunt.</p>}
              </div>
            </Card>
          );
        })}
        {!vergadering.agendapunten.length && <EmptyState icon={ClipboardList} text="Nog geen agendapunten — voeg er een toe om te starten." />}
      </div>

      {actieVoorPunt && <ActieForm punt={actieVoorPunt} onSave={data => addActie(actieVoorPunt, data)} onClose={() => setActieVoorPunt(null)} />}
      {showAddPunt && (
        <PromptModal title="Agendapunt toevoegen" label="Titel van het agendapunt" placeholder="bv. Financieel overzicht"
          onSave={addPunt} onCancel={() => setShowAddPunt(false)} />
      )}
      {delPuntId != null && (
        <ConfirmModal message="Dit agendapunt en de bijbehorende notulen verwijderen?" onConfirm={() => { removePunt(delPuntId); setDelPuntId(null); }} onCancel={() => setDelPuntId(null)} />
      )}
      {delActieId != null && (
        <ConfirmModal message="Dit actiepunt verwijderen?" onConfirm={() => { removeActie(delActieId); setDelActieId(null); }} onCancel={() => setDelActieId(null)} />
      )}
    </div>
  );
}

function ActieForm({ punt, onSave, onClose }) {
  const [f, setF] = useState({ omschrijving: '', wie: '', deadline: '' });
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  return (
    <Modal title={`Actiepunt bij "${punt.titel}"`} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Omschrijving"><input className={inputCls} style={inputStyle} value={f.omschrijving} onChange={e => upd('omschrijving', e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Wie"><input className={inputCls} style={inputStyle} value={f.wie} onChange={e => upd('wie', e.target.value)} /></Field>
          <Field label="Deadline"><input type="date" className={inputCls} style={inputStyle} value={f.deadline} onChange={e => upd('deadline', e.target.value)} /></Field>
        </div>
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(f)} disabled={!f.omschrijving}>Toevoegen</Btn>
      </div>
    </Modal>
  );
}

function ActielijstView({ actielijst, setActielijst, logoHoogteCm, onBack, readOnly }) {
  const [status, setStatus] = useState('open');
  const [q, setQ] = useState('');
  const [delId, setDelId] = useState(null);

  const filtered = actielijst.filter(a => {
    if (status !== 'alle' && a.status !== status) return false;
    if (q) {
      const hay = `${a.omschrijving} ${a.wie || ''} ${a.vergaderingTitel || ''}`.toLowerCase();
      if (!hay.includes(q.toLowerCase())) return false;
    }
    return true;
  }).sort((a, b) => (a.deadline || '9999').localeCompare(b.deadline || '9999'));

  function toggle(id) {
    setActielijst(actielijst.map(a => a.id === id ? { ...a, status: a.status === 'klaar' ? 'open' : 'klaar' } : a));
  }
  function remove(id) {
    setActielijst(actielijst.filter(a => a.id !== id));
  }
  function exportWord() {
    downloadWordDoc({ titel: 'Actielijst', filename: 'BladelsCreatief_Actielijst.doc', bodyHtml: actielijstWordHtml(filtered, logoHoogteCm) });
  }

  return (
    <div className="space-y-4">
      <button onClick={onBack} className="flex items-center gap-1 text-sm" style={{ color: C.inkSoft }}><ArrowLeft size={15} /> Terug naar vergaderingen</button>

      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2">
          <select value={status} onChange={e => setStatus(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="open">Open</option>
            <option value="klaar">Klaar</option>
            <option value="alle">Alle</option>
          </select>
          <div className="relative">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: C.inkSoft }} />
            <input value={q} onChange={e => setQ(e.target.value)} placeholder="Zoek…" className={`${inputCls} pl-8`} style={{ ...inputStyle, width: 190 }} />
          </div>
        </div>
        <Btn tone="sage" icon={FileText} onClick={exportWord} disabled={!filtered.length}>Als Word</Btn>
      </div>

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              {['', 'Actiepunt', 'Vergadering', 'Agendapunt', 'Wie', 'Deadline', ''].map(h => (
                <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map(a => (
              <tr key={a.id} className="border-b last:border-0 hover:bg-black/[0.02]" style={{ borderColor: C.border }}>
                <td className="px-3 py-2">
                  <button onClick={() => !readOnly && toggle(a.id)} disabled={readOnly} style={{ color: a.status === 'klaar' ? C.sageDeep : C.inkSoft }}>
                    {a.status === 'klaar' ? <CheckSquare size={16} /> : <Square size={16} />}
                  </button>
                </td>
                <td className="px-3 py-2" style={{ textDecoration: a.status === 'klaar' ? 'line-through' : 'none', color: a.status === 'klaar' ? C.inkSoft : C.ink }}>{a.omschrijving}</td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>{a.vergaderingTitel || '—'}</td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>{a.agendapuntTitel || '—'}</td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>{a.wie || '—'}</td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>{a.deadline ? fmtDate(a.deadline) : '—'}</td>
                <td className="px-3 py-2">{!readOnly && <button onClick={() => setDelId(a.id)} style={{ color: C.rose }}><Trash2 size={14} /></button>}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {!filtered.length && <EmptyState icon={CheckSquare} text="Geen actiepunten gevonden met deze filters." />}
      </Card>
      {delId != null && (
        <ConfirmModal message="Dit actiepunt verwijderen?" onConfirm={() => { remove(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
    </div>
  );
}

/* =========================================================================
   WORKSHOPS — generieke workshopadministratie (incl. jeugdatelier-blokken)
========================================================================= */
function WorkshopsTab({ members, workshops, setWorkshops, inschrijvingen, setInschrijvingen, tx, setTx, accounts, workshopSoorten, logoHoogteCm, workshopSortering, setWorkshopSortering, readOnly, initialQuery, onTrash, onLog }) {
  const [view, setView] = useState('lijst');
  const [selectedId, setSelectedId] = useState(null);
  const selected = workshops.find(w => w.id === selectedId);

  if (view === 'detail' && selected) {
    return <WorkshopDetail key={selected.id} members={members} workshop={selected} workshops={workshops} setWorkshops={setWorkshops}
      inschrijvingen={inschrijvingen} setInschrijvingen={setInschrijvingen} tx={tx} setTx={setTx} accounts={accounts} workshopSoorten={workshopSoorten}
      readOnly={readOnly} onBack={() => setView('lijst')} onTrash={onTrash} onLog={onLog} />;
  }
  return (
    <WorkshopList workshops={workshops} setWorkshops={setWorkshops} inschrijvingen={inschrijvingen} setInschrijvingen={setInschrijvingen}
      workshopSoorten={workshopSoorten} logoHoogteCm={logoHoogteCm} workshopSortering={workshopSortering} setWorkshopSortering={setWorkshopSortering}
      readOnly={readOnly} initialQuery={initialQuery} onOpen={id => { setSelectedId(id); setView('detail'); }}
      onTrash={onTrash} onLog={onLog} />
  );
}

function WorkshopList({ workshops, setWorkshops, inschrijvingen, setInschrijvingen, workshopSoorten, logoHoogteCm, workshopSortering, setWorkshopSortering, readOnly, initialQuery, onOpen, onTrash, onLog }) {
  const [q, setQ] = useState(initialQuery || '');
  const [soort, setSoort] = useState('alle');
  const [status, setStatus] = useState('open');
  const [sortField, setSortField] = useState(workshopSortering.veld);
  const [sortDir, setSortDir] = useState(workshopSortering.richting);
  const [showNew, setShowNew] = useState(false);
  const [delId, setDelId] = useState(null);
  const [geselecteerd, setGeselecteerd] = useState([]);
  const [showOverzicht, setShowOverzicht] = useState(false);
  function toggleSelectie(id) {
    setGeselecteerd(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
  }

  function toggleSort(field) {
    if (sortField === field) {
      const nieuweRichting = sortDir === 'asc' ? 'desc' : 'asc';
      setSortDir(nieuweRichting);
      setWorkshopSortering({ veld: field, richting: nieuweRichting });
    } else {
      setSortField(field);
      setSortDir('asc');
      setWorkshopSortering({ veld: field, richting: 'asc' });
    }
  }
  function sortIcon(field) {
    if (sortField !== field) return '';
    return sortDir === 'asc' ? ' \u25B2' : ' \u25BC';
  }

  const soorten = Array.from(new Set([...workshopSoorten, ...workshops.map(w => w.soort)])).sort();
  const filtered = workshops.filter(w => {
    if (status !== 'alle' && w.status !== status) return false;
    if (soort !== 'alle' && w.soort !== soort) return false;
    if (q && !`${w.titel} ${w.soort}`.toLowerCase().includes(q.toLowerCase())) return false;
    return true;
  }).sort((a, b) => {
    let cmp;
    if (sortField === 'soort') cmp = a.soort.localeCompare(b.soort) || a.titel.localeCompare(b.titel);
    else if (sortField === 'titel') cmp = a.titel.localeCompare(b.titel);
    else cmp = (a.datums[0] || '').localeCompare(b.datums[0] || '') || a.titel.localeCompare(b.titel);
    return sortDir === 'asc' ? cmp : -cmp;
  });

  function telling(w) {
    const items = inschrijvingen.filter(i => i.workshopId === w.id);
    return {
      ingeschreven: items.filter(i => i.status === 'ingeschreven').length,
      wachtlijst: items.filter(i => i.status === 'wachtlijst').length,
    };
  }

  function createWorkshop(f) {
    const id = uid(workshops);
    setWorkshops([...workshops, { id, ...f }]);
    onLog(`Workshop aangemaakt: ${f.titel} (${f.soort})`, 'workshops');
    setShowNew(false);
    onOpen(id);
  }
  function removeWorkshop(id) {
    const w = workshops.find(x => x.id === id);
    const gekoppeld = inschrijvingen.filter(i => i.workshopId === id);
    setWorkshops(workshops.filter(x => x.id !== id));
    setInschrijvingen(inschrijvingen.filter(i => i.workshopId !== id));
    if (w) {
      onTrash('workshop', w);
      gekoppeld.forEach(i => onTrash('inschrijving', i));
      onLog(`Workshop verwijderd: ${w.titel} (incl. ${gekoppeld.length} inschrijvingen)`, 'workshops');
    }
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2">
          <div className="relative">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: C.inkSoft }} />
            <input value={q} onChange={e => setQ(e.target.value)} placeholder="Zoek workshop…" className={`${inputCls} pl-8`} style={{ ...inputStyle, width: 200 }} />
          </div>
          <select value={status} onChange={e => setStatus(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="open">Open</option>
            <option value="gesloten">Gesloten</option>
            <option value="afgerond">Afgerond</option>
            <option value="alle">Alle statussen</option>
          </select>
          <select value={soort} onChange={e => setSoort(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Alle soorten</option>
            {soorten.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <div className="flex items-center gap-2">
          <Btn icon={FileText} tone="outline" onClick={() => setShowOverzicht(true)} disabled={!geselecteerd.length}>Overzicht maken{geselecteerd.length ? ` (${geselecteerd.length})` : ''}</Btn>
          {!readOnly && <Btn icon={Plus} onClick={() => setShowNew(true)}>Nieuwe workshop</Btn>}
        </div>
      </div>

      <div className="flex items-center gap-1.5 flex-wrap">
        <span className="text-xs" style={{ color: C.inkSoft }}>Sorteer op:</span>
        {[
          { key: 'datum', label: 'Startdatum' },
          { key: 'soort', label: 'Soort' },
          { key: 'titel', label: 'Titel' },
        ].map(opt => (
          <button key={opt.key} onClick={() => toggleSort(opt.key)}
            className="px-2.5 py-1 rounded-full text-xs border"
            style={{ background: sortField === opt.key ? C.clay : 'transparent', color: sortField === opt.key ? '#fff' : C.ink, borderColor: sortField === opt.key ? C.clay : C.border }}>
            {opt.label}{sortIcon(opt.key)}
          </button>
        ))}
      </div>

      <div className="grid sm:grid-cols-2 gap-3">
        {filtered.map(w => {
          const t = telling(w);
          return (
            <Card key={w.id} className="p-4 cursor-pointer hover:shadow-md transition-shadow" onClick={() => onOpen(w.id)}>
              <div className="flex items-start justify-between">
                <div className="flex items-start gap-2">
                  <input type="checkbox" className="mt-1" checked={geselecteerd.includes(w.id)} onClick={e => e.stopPropagation()} onChange={() => toggleSelectie(w.id)} />
                  <div>
                    <h3 className="font-semibold" style={{ color: C.ink }}>{w.titel}</h3>
                    <p className="text-xs" style={{ color: C.inkSoft }}>
                      {w.type === 'reeks' ? `Reeks · ${w.datums.length} data` : (w.datums[0] ? fmtDate(w.datums[0]) : 'Datum nog niet ingepland')}
                      {w.dagdeel ? ` · ${w.dagdeel}` : ''}
                    </p>
                  </div>
                </div>
                {!readOnly && <button onClick={e => { e.stopPropagation(); setDelId(w.id); }} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={14} /></button>}
              </div>
              <div className="flex gap-1.5 mt-2.5 flex-wrap">
                <Badge tone="clay">{w.soort}</Badge>
                <Badge tone="sage">{t.ingeschreven}{w.maxDeelnemers ? ` / ${w.maxDeelnemers}` : ''} ingeschreven</Badge>
                {t.wachtlijst > 0 && <Badge tone="ochre">{t.wachtlijst} wachtlijst</Badge>}
                <Badge tone="muted">{w.status}</Badge>
              </div>
            </Card>
          );
        })}
        {!filtered.length && (
          <div className="sm:col-span-2"><EmptyState icon={Palette} text="Geen workshops gevonden met deze filters." /></div>
        )}
      </div>

      {showNew && <NewWorkshopModal workshopSoorten={workshopSoorten} onSave={createWorkshop} onClose={() => setShowNew(false)} />}
      {delId != null && (
        <ConfirmModal message="Deze workshop verwijderen? De gekoppelde inschrijvingen worden ook verwijderd (en zijn terug te vinden in de prullenbak)."
          onConfirm={() => { removeWorkshop(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
      {showOverzicht && (
        <Modal title="Workshopoverzicht exporteren" onClose={() => setShowOverzicht(false)}>
          <p className="text-sm mb-4" style={{ color: C.inkSoft }}>
            {geselecteerd.length} workshop(s) geselecteerd. Het overzicht bevat per workshop de gegevens en de volledige deelnemerslijst.
          </p>
          <div className="flex gap-2">
            <Btn tone="outline" icon={Printer} onClick={() => {
              const gekozen = filtered.filter(w => geselecteerd.includes(w.id));
              downloadPrintableHtmlVrij({ titel: 'Workshopoverzicht', bodyHtml: workshopsOverzichtKern(gekozen, inschrijvingen) });
              setShowOverzicht(false);
            }}>Als PDF</Btn>
            <Btn tone="sage" icon={FileText} onClick={() => {
              const gekozen = filtered.filter(w => geselecteerd.includes(w.id));
              const body = `${logoImgTag(logoHoogteCm)}<div class="bar"></div>${workshopsOverzichtKern(gekozen, inschrijvingen)}`;
              downloadWordDoc({ titel: 'Workshopoverzicht', filename: 'BladelsCreatief_Workshopoverzicht.doc', bodyHtml: body });
              setShowOverzicht(false);
            }}>Als Word</Btn>
          </div>
          <div className="flex justify-end mt-5">
            <Btn tone="ghost" onClick={() => setShowOverzicht(false)}>Sluiten</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

function NewWorkshopModal({ workshopSoorten, onSave, onClose }) {
  const [f, setF] = useState({ titel: '', soort: 'Jeugdatelier', type: 'eenmalig', datums: [new Date().toISOString().slice(0, 10)], dagdeel: '', locatie: '', bedrag: '', maxDeelnemers: '', status: 'open' });
  const [soortAnders, setSoortAnders] = useState(false);
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  function kiesSoort(v) {
    if (v === '__anders__') { setSoortAnders(true); upd('soort', ''); }
    else { setSoortAnders(false); upd('soort', v); }
  }
  function setDatum(i, v) { const d = [...f.datums]; d[i] = v; upd('datums', d); }
  function addDatum() { upd('datums', [...f.datums, '']); }
  function removeDatum(i) { upd('datums', f.datums.filter((_, idx) => idx !== i)); }

  return (
    <Modal title="Nieuwe workshop" onClose={onClose} wide>
      <div className="space-y-3">
        <Field label="Titel"><input className={inputCls} style={inputStyle} value={f.titel} onChange={e => upd('titel', e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Soort">
            <select className={inputCls} style={inputStyle} value={soortAnders ? '__anders__' : f.soort} onChange={e => kiesSoort(e.target.value)}>
              {workshopSoorten.map(s => <option key={s} value={s}>{s}</option>)}
              <option value="__anders__">Andere soort...</option>
            </select>
            {soortAnders && (
              <input className={inputCls} style={{ ...inputStyle, marginTop: 6 }} placeholder="Naam van de nieuwe soort" value={f.soort} onChange={e => upd('soort', e.target.value)} />
            )}
          </Field>
          <Field label="Type">
            <select className={inputCls} style={inputStyle} value={f.type} onChange={e => upd('type', e.target.value)}>
              <option value="eenmalig">Eenmalig</option>
              <option value="reeks">Reeks (meerdere data)</option>
            </select>
          </Field>
        </div>
        <div>
          <span className="block text-xs font-medium mb-1" style={{ color: C.inkSoft }}>{f.type === 'reeks' ? 'Data' : 'Datum'}</span>
          {f.datums.map((d, i) => (
            <div key={i} className="flex items-center gap-2 mb-1.5">
              <input type="date" className={inputCls} style={inputStyle} value={d} onChange={e => setDatum(i, e.target.value)} />
              {f.type === 'reeks' && f.datums.length > 1 && <button type="button" onClick={() => removeDatum(i)} style={{ color: C.rose }}><Trash2 size={14} /></button>}
            </div>
          ))}
          {f.type === 'reeks' && <button type="button" onClick={addDatum} className="text-xs underline" style={{ color: C.clay }}>+ datum toevoegen</button>}
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Dagdeel/tijd"><input className={inputCls} style={inputStyle} placeholder="bv. wo 19:00-21:00" value={f.dagdeel} onChange={e => upd('dagdeel', e.target.value)} /></Field>
          <Field label="Locatie"><input className={inputCls} style={inputStyle} value={f.locatie} onChange={e => upd('locatie', e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Bedrag per deelnemer (optioneel)"><input type="number" step="0.01" className={inputCls} style={inputStyle} value={f.bedrag} onChange={e => upd('bedrag', e.target.value)} /></Field>
          <Field label="Max. aantal deelnemers (optioneel)"><input type="number" className={inputCls} style={inputStyle} value={f.maxDeelnemers} onChange={e => upd('maxDeelnemers', e.target.value)} /></Field>
        </div>
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave({ ...f, bedrag: f.bedrag === '' ? null : Number(f.bedrag), maxDeelnemers: f.maxDeelnemers === '' ? null : Number(f.maxDeelnemers), datums: f.datums.filter(Boolean) })} disabled={!f.titel || !f.soort}>Aanmaken</Btn>
      </div>
    </Modal>
  );
}

function WorkshopDetail({ members, workshop, workshops, setWorkshops, inschrijvingen, setInschrijvingen, tx, setTx, accounts, workshopSoorten, readOnly, onBack, onTrash, onLog }) {
  const [showInschrijving, setShowInschrijving] = useState(false);
  const [editingInschrijving, setEditingInschrijving] = useState(null);
  const [betaalFor, setBetaalFor] = useState(null);
  const [delId, setDelId] = useState(null);
  const [soortAnders, setSoortAnders] = useState(!workshopSoorten.includes(workshop.soort));

  function update(patch) {
    setWorkshops(workshops.map(w => w.id === workshop.id ? { ...w, ...patch } : w));
  }
  const mijnInschrijvingen = inschrijvingen.filter(i => i.workshopId === workshop.id)
    .sort((a, b) => (a.status === b.status ? 0 : a.status === 'wachtlijst' ? 1 : -1) || (a.naam || '').localeCompare(b.naam || ''));
  const aantalIngeschreven = mijnInschrijvingen.filter(i => i.status === 'ingeschreven').length;
  const vol = workshop.maxDeelnemers != null && aantalIngeschreven >= workshop.maxDeelnemers;

  function saveInschrijving(data) {
    if (data.id) {
      // Bewerken: contactgegevens/groep/etc. bijwerken, status/betaling blijven ongewijzigd.
      setInschrijvingen(inschrijvingen.map(i => i.id === data.id ? { ...i, ...data } : i));
      onLog(`Inschrijving bewerkt bij ${workshop.titel}: ${data.naam}`, 'workshops');
    } else {
      const status = (workshop.maxDeelnemers != null && aantalIngeschreven >= workshop.maxDeelnemers) ? 'wachtlijst' : 'ingeschreven';
      const nieuw = { id: uid(inschrijvingen), workshopId: workshop.id, ...data, status };
      setInschrijvingen([...inschrijvingen, nieuw]);
      onLog(`Inschrijving toegevoegd bij ${workshop.titel}: ${data.naam}${status === 'wachtlijst' ? ' (wachtlijst)' : ''}`, 'workshops');
    }
    setShowInschrijving(false);
    setEditingInschrijving(null);
  }
  function removeInschrijving(id) {
    const i = inschrijvingen.find(x => x.id === id);
    setInschrijvingen(inschrijvingen.filter(x => x.id !== id));
    if (i) { onTrash('inschrijving', i); onLog(`Inschrijving verwijderd bij ${workshop.titel}: ${i.naam}`, 'workshops'); }
  }
  function promoveer(id) {
    const i = inschrijvingen.find(x => x.id === id);
    setInschrijvingen(inschrijvingen.map(x => x.id === id ? { ...x, status: 'ingeschreven' } : x));
    if (i) onLog(`Van wachtlijst naar ingeschreven: ${i.naam} (${workshop.titel})`, 'workshops');
  }
  function naarWachtlijst(id) {
    setInschrijvingen(inschrijvingen.map(i => i.id === id ? { ...i, status: 'wachtlijst' } : i));
  }
  function markeerBetaald(inschrijving, data) {
    setInschrijvingen(inschrijvingen.map(i => i.id === inschrijving.id ? { ...i, betaald: true, bedrag: Number(data.bedrag) } : i));
    onLog(`Workshopbetaling ontvangen: ${inschrijving.naam} (${euro(data.bedrag)})`, 'workshops');
    if (data.boeken) {
      const acc = accounts.find(a => a.code === data.grootboekCode);
      const nieuw = {
        id: uid(tx), rekening: data.rekening, jaar: new Date(data.datum).getFullYear(), maand: new Date(data.datum).getMonth() + 1,
        datum: data.datum, grootboek_code: data.grootboekCode, grootboek_naam: acc ? acc.naam : '',
        bedrag: Math.abs(Number(data.bedrag)), omschrijving: `${workshop.titel} - ${inschrijving.naam}`,
        gekoppeldType: 'workshop', gekoppeldRef: { inschrijvingId: inschrijving.id },
      };
      setTx([...tx, nieuw]);
      onLog(`Boeking toegevoegd voor workshop: ${inschrijving.naam} (${euro(nieuw.bedrag)})`, 'financien');
    }
    setBetaalFor(null);
  }
  const [openZonderKoppeling, setOpenZonderKoppeling] = useState(null);
  function gekoppeldeBoeking(inschrijving) {
    return tx.find(t => t.gekoppeldType === 'workshop' && t.gekoppeldRef.inschrijvingId === inschrijving.id);
  }
  function markeerOnbetaald(inschrijving) {
    const boeking = gekoppeldeBoeking(inschrijving);
    if (boeking) { setOpenZonderKoppeling(inschrijving); return; }
    doeMarkeerOnbetaald(inschrijving);
  }
  function doeMarkeerOnbetaald(inschrijving) {
    setInschrijvingen(inschrijvingen.map(i => i.id === inschrijving.id ? { ...i, betaald: false } : i));
    onLog(`Workshopbetaling teruggezet naar open: ${inschrijving.naam}`, 'workshops');
    const boeking = gekoppeldeBoeking(inschrijving);
    if (boeking) setTx(tx.map(t => t.id === boeking.id ? { ...t, gekoppeldType: null, gekoppeldRef: null } : t));
    setOpenZonderKoppeling(null);
  }

  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const aoa = [[workshop.titel], [`${workshop.soort} · ${workshop.dagdeel || ''}`], [],
      ['Naam', 'Groep', 'Herkomst', 'Contact', 'Contact ouder 2', 'Inschrijfdatum', 'Status', 'Betaald', 'Bedrag'],
      ...mijnInschrijvingen.map(i => [
        i.naam, i.groep || '', i.herkomst, i.email || i.telefoon || '', i.email2 || i.telefoon2 || '',
        i.datumInschrijving ? fmtDate(i.datumInschrijving) : '', i.status, i.betaald ? 'ja' : 'nee', i.bedrag ? euro(i.bedrag) : '',
      ])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Deelnemers');
    XLSX.writeFile(wb, `BladelsCreatief_${workshop.titel.replace(/\s+/g, '_')}.xlsx`);
  }
  function exportPdf() {
    downloadPrintableHtml({
      titel: workshop.titel,
      subtitel: `${workshop.soort} · ${aantalIngeschreven} ingeschreven${workshop.maxDeelnemers != null ? ` / max ${workshop.maxDeelnemers}` : ''}`,
      header: ['Naam', 'Groep', 'Herkomst', 'Contact', 'Inschrijfdatum', 'Status', 'Betaald'],
      rows: mijnInschrijvingen.map(i => [
        i.naam, i.groep || '—', i.herkomst === 'lid' ? 'Lid' : 'Extern', i.email || i.telefoon || '—',
        i.datumInschrijving ? fmtDate(i.datumInschrijving) : '—', i.status === 'wachtlijst' ? 'Wachtlijst' : 'Ingeschreven', i.betaald ? 'Ja' : 'Nee',
      ]),
    });
  }

  return (
    <div className="space-y-4">
      <button onClick={onBack} className="flex items-center gap-1 text-sm" style={{ color: C.inkSoft }}><ArrowLeft size={15} /> Terug naar workshops</button>

      <Card className="p-4">
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label="Titel"><DebouncedField disabled={readOnly} value={workshop.titel} onCommit={v => update({ titel: v })} /></Field>
          <Field label="Soort">
            <select disabled={readOnly} className={inputCls} style={inputStyle}
              value={soortAnders ? '__anders__' : workshop.soort}
              onChange={e => {
                if (e.target.value === '__anders__') { setSoortAnders(true); update({ soort: '' }); }
                else { setSoortAnders(false); update({ soort: e.target.value }); }
              }}>
              {Array.from(new Set([...workshopSoorten, workshop.soort])).map(s => <option key={s} value={s}>{s}</option>)}
              <option value="__anders__">Andere soort...</option>
            </select>
            {soortAnders && (
              <input disabled={readOnly} className={inputCls} style={{ ...inputStyle, marginTop: 6 }} placeholder="Naam van de nieuwe soort" value={workshop.soort} onChange={e => update({ soort: e.target.value })} />
            )}
          </Field>
          <Field label="Dagdeel/tijd"><DebouncedField disabled={readOnly} value={workshop.dagdeel || ''} onCommit={v => update({ dagdeel: v })} /></Field>
          <Field label="Locatie"><DebouncedField disabled={readOnly} value={workshop.locatie || ''} onCommit={v => update({ locatie: v })} /></Field>
          <Field label="Bedrag per deelnemer">
            <DebouncedField type="number" disabled={readOnly} value={workshop.bedrag ?? ''} onCommit={v => update({ bedrag: v === '' ? null : Number(v) })} />
          </Field>
          <Field label="Max. aantal deelnemers">
            <DebouncedField type="number" disabled={readOnly} value={workshop.maxDeelnemers ?? ''} onCommit={v => update({ maxDeelnemers: v === '' ? null : Number(v) })} />
          </Field>
          <Field label="Status">
            <select disabled={readOnly} className={inputCls} style={inputStyle} value={workshop.status} onChange={e => update({ status: e.target.value })}>
              <option value="open">Open</option>
              <option value="gesloten">Gesloten (vol / inschrijving dicht)</option>
              <option value="afgerond">Afgerond</option>
            </select>
          </Field>
        </div>

        <div className="mt-3">
          <span className="block text-xs font-medium mb-1" style={{ color: C.inkSoft }}>{workshop.type === 'reeks' ? 'Data (reeks)' : 'Datum'}</span>
          <div className="flex flex-wrap gap-2">
            {workshop.datums.map((d, i) => (
              <div key={i} className="flex items-center gap-1">
                <input disabled={readOnly} type="date" className="rounded border px-2 py-1 text-sm" style={inputStyle} value={d}
                  onChange={e => { const nd = [...workshop.datums]; nd[i] = e.target.value; update({ datums: nd }); }} />
                {!readOnly && workshop.type === 'reeks' && <button onClick={() => update({ datums: workshop.datums.filter((_, idx) => idx !== i) })} style={{ color: C.rose }}><Trash2 size={13} /></button>}
              </div>
            ))}
            {!readOnly && workshop.type === 'reeks' && (
              <button onClick={() => update({ datums: [...workshop.datums, ''] })} className="text-xs underline" style={{ color: C.clay }}>+ datum</button>
            )}
          </div>
        </div>

        <div className="flex justify-end gap-2 mt-4">
          <Btn tone="outline" icon={Printer} onClick={exportPdf}>Als PDF</Btn>
          <Btn tone="sage" icon={Download} onClick={exportExcel}>Als Excel</Btn>
        </div>
      </Card>

      <div className="flex items-center justify-between flex-wrap gap-2">
        <h3 className="font-semibold flex items-center gap-1.5" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>
          Deelnemers <Badge tone={vol ? 'ochre' : 'sage'}>{aantalIngeschreven}{workshop.maxDeelnemers != null ? ` / ${workshop.maxDeelnemers}` : ''}</Badge>
        </h3>
        {!readOnly && <Btn tone="outline" size="sm" icon={Plus} onClick={() => setShowInschrijving(true)}>Inschrijving toevoegen</Btn>}
      </div>
      {vol && <p className="text-xs" style={{ color: C.ochre }}>Deze workshop zit vol — nieuwe inschrijvingen komen automatisch op de wachtlijst.</p>}

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              {['Naam', 'Herkomst', 'Contact', 'Inschrijfdatum', 'Status', 'Betaald', ''].map(h => <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {mijnInschrijvingen.map(i => (
              <tr key={i.id} className="border-b last:border-0" style={{ borderColor: C.border }}>
                <td className="px-3 py-2 font-medium">
                  {i.naam}{i.leeftijd ? <span className="text-xs" style={{ color: C.inkSoft }}> ({i.leeftijd} jr)</span> : ''}
                  {i.groep ? <span className="block text-xs font-normal" style={{ color: C.inkSoft }}>Groep {i.groep}</span> : ''}
                </td>
                <td className="px-3 py-2"><Badge tone={i.herkomst === 'lid' ? 'clay' : 'muted'}>{i.herkomst === 'lid' ? 'Lid' : 'Extern'}</Badge></td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>
                  <div>{[i.email, i.telefoon].filter(Boolean).join(' · ') || '—'}</div>
                  {(i.email2 || i.telefoon2) && <div className="text-xs">{[i.email2, i.telefoon2].filter(Boolean).join(' · ')} (ouder 2)</div>}
                </td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>{i.datumInschrijving ? fmtDate(i.datumInschrijving) : '—'}</td>
                <td className="px-3 py-2"><Badge tone={i.status === 'wachtlijst' ? 'ochre' : 'sage'}>{i.status === 'wachtlijst' ? 'Wachtlijst' : 'Ingeschreven'}</Badge></td>
                <td className="px-3 py-2">
                  {!readOnly ? (
                    <button onClick={() => i.betaald ? markeerOnbetaald(i) : setBetaalFor(i)}>
                      <Badge tone={i.betaald ? 'sage' : 'rose'}>{i.betaald ? 'Betaald (klik: open)' : 'Markeer betaald'}</Badge>
                    </button>
                  ) : <Badge tone={i.betaald ? 'sage' : 'rose'}>{i.betaald ? 'Betaald' : 'Open'}</Badge>}
                </td>
                <td className="px-3 py-2">
                  {!readOnly && (
                    <div className="flex gap-2 justify-end items-center">
                      {i.status === 'wachtlijst'
                        ? <button onClick={() => promoveer(i.id)} className="text-xs underline" style={{ color: C.sageDeep }}>naar ingeschreven</button>
                        : <button onClick={() => naarWachtlijst(i.id)} className="text-xs underline" style={{ color: C.inkSoft }}>naar wachtlijst</button>}
                      <button onClick={() => setEditingInschrijving(i)} style={{ color: C.inkSoft }}><Pencil size={13} /></button>
                      <button onClick={() => setDelId(i.id)} style={{ color: C.rose }}><Trash2 size={13} /></button>
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!mijnInschrijvingen.length && <EmptyState icon={Users} text="Nog geen inschrijvingen voor deze workshop." />}
      </Card>

      {(showInschrijving || editingInschrijving) && (
        <InschrijvingForm members={members} inschrijvingen={inschrijvingen} workshopSoort={workshop.soort} item={editingInschrijving}
          onSave={saveInschrijving} onClose={() => { setShowInschrijving(false); setEditingInschrijving(null); }} />
      )}
      {betaalFor && <BetaalWorkshopForm inschrijving={betaalFor} accounts={accounts} onSave={data => markeerBetaald(betaalFor, data)} onClose={() => setBetaalFor(null)} />}
      {openZonderKoppeling && (
        <ConfirmModal title="Betaling op open zetten"
          message={`Dit is gekoppeld aan een boeking in Financiën (${gekoppeldeBoeking(openZonderKoppeling)?.omschrijving || ''}). Doorgaan zet de betaling op open en ontkoppelt de boeking — de boeking zelf blijft gewoon bestaan.`}
          onConfirm={() => doeMarkeerOnbetaald(openZonderKoppeling)} onCancel={() => setOpenZonderKoppeling(null)} />
      )}
      {delId != null && (
        <ConfirmModal message="Deze inschrijving verwijderen?" onConfirm={() => { removeInschrijving(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
    </div>
  );
}

function InschrijvingForm({ members, inschrijvingen, workshopSoort, item, onSave, onClose }) {
  const [herkomst, setHerkomst] = useState(item ? item.herkomst : 'extern');
  const [lidId, setLidId] = useState(item && item.lidId ? String(item.lidId) : '');
  const [f, setF] = useState(item ? {
    naam: item.naam || '', email: item.email || '', email2: item.email2 || '', telefoon: item.telefoon || '', telefoon2: item.telefoon2 || '',
    leeftijd: item.leeftijd ?? '', groep: item.groep || '', datumInschrijving: item.datumInschrijving || new Date().toISOString().slice(0, 10),
  } : {
    naam: '', email: '', email2: '', telefoon: '', telefoon2: '', leeftijd: '', groep: '',
    datumInschrijving: new Date().toISOString().slice(0, 10),
  });
  const isJeugd = workshopSoort === 'Jeugdatelier';
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }

  // Eerdere externe deelnemers (over alle workshops heen) — voor suggesties en slim aanvullen op naam (alleen bij nieuwe inschrijvingen).
  const eerdereNamen = Array.from(new Set(inschrijvingen.filter(i => i.herkomst === 'extern' && i.naam).map(i => i.naam))).sort();
  const heeftEerdereMatch = !item && herkomst === 'extern' && f.naam.trim() && eerdereNamen.some(n => n.toLowerCase() === f.naam.trim().toLowerCase());

  function naamChange(v) {
    if (item) { upd('naam', v); return; }
    const match = inschrijvingen
      .filter(i => i.herkomst === 'extern' && i.naam && i.naam.trim().toLowerCase() === v.trim().toLowerCase())
      .sort((a, b) => (b.datumInschrijving || '').localeCompare(a.datumInschrijving || ''))[0];
    if (match) {
      setF(s => ({
        ...s, naam: v,
        email: match.email || s.email, email2: match.email2 || s.email2,
        telefoon: match.telefoon || s.telefoon, telefoon2: match.telefoon2 || s.telefoon2,
        leeftijd: match.leeftijd ?? s.leeftijd, groep: match.groep || s.groep,
      }));
    } else {
      upd('naam', v);
    }
  }

  function kiesLid(id) {
    setLidId(id);
    const m = members.find(x => String(x.id) === String(id));
    if (m) setF(s => ({ ...s, naam: fullName(m), email: m.email || '', telefoon: m.telefoon || '' }));
  }

  function submit() {
    onSave({
      ...(item ? { id: item.id, status: item.status, betaald: item.betaald, bedrag: item.bedrag } : { betaald: false, bedrag: null }),
      herkomst, lidId: herkomst === 'lid' ? Number(lidId) : null,
      naam: f.naam, email: f.email, email2: isJeugd ? f.email2 : '', telefoon: f.telefoon, telefoon2: isJeugd ? f.telefoon2 : '',
      leeftijd: f.leeftijd === '' ? null : Number(f.leeftijd), groep: isJeugd ? f.groep : '',
      datumInschrijving: f.datumInschrijving, notitie: '',
    });
  }
  return (
    <Modal title={item ? 'Inschrijving bewerken' : 'Inschrijving toevoegen'} onClose={onClose}>
      <div className="space-y-3">
        <div className="flex gap-1">
          {[['extern', 'Extern (geen lid)'], ['lid', 'Bestaand lid']].map(([id, label]) => (
            <button key={id} onClick={() => setHerkomst(id)} className="px-3 py-1.5 rounded-lg text-sm font-medium border"
              style={{ background: herkomst === id ? C.clay : 'transparent', color: herkomst === id ? '#fff' : C.ink, borderColor: herkomst === id ? C.clay : C.border }}>
              {label}
            </button>
          ))}
        </div>
        {herkomst === 'lid' && (
          <Field label="Kies lid">
            <select className={inputCls} style={inputStyle} value={lidId} onChange={e => kiesLid(e.target.value)}>
              <option value="">— kies —</option>
              {members.slice().sort((a, b) => fullName(a).localeCompare(fullName(b))).map(m => <option key={m.id} value={m.id}>{fullName(m)}</option>)}
            </select>
          </Field>
        )}
        <Field label="Naam">
          <input list="eerdere-deelnemers" className={inputCls} style={inputStyle} value={f.naam} onChange={e => naamChange(e.target.value)} />
          <datalist id="eerdere-deelnemers">{eerdereNamen.map(n => <option key={n} value={n} />)}</datalist>
        </Field>
        {heeftEerdereMatch && (
          <p className="text-xs" style={{ color: C.sageDeep }}>Eerdere gegevens van deze deelnemer automatisch aangevuld — controleer en pas aan waar nodig.</p>
        )}
        <div className="grid grid-cols-2 gap-3">
          <Field label={isJeugd ? 'E-mail ouder 1' : 'E-mail'}><input className={inputCls} style={inputStyle} value={f.email} onChange={e => upd('email', e.target.value)} /></Field>
          <Field label={isJeugd ? 'Telefoonnr. ouder 1' : 'Telefoon'}><input className={inputCls} style={inputStyle} value={f.telefoon} onChange={e => upd('telefoon', e.target.value)} /></Field>
        </div>
        {isJeugd && (
          <div className="grid grid-cols-2 gap-3">
            <Field label="E-mail ouder 2"><input className={inputCls} style={inputStyle} value={f.email2} onChange={e => upd('email2', e.target.value)} /></Field>
            <Field label="Telefoonnr. ouder 2"><input className={inputCls} style={inputStyle} value={f.telefoon2} onChange={e => upd('telefoon2', e.target.value)} /></Field>
          </div>
        )}
        <div className="grid grid-cols-2 gap-3">
          <Field label="Leeftijd (optioneel)"><input type="number" className={inputCls} style={inputStyle} value={f.leeftijd} onChange={e => upd('leeftijd', e.target.value)} /></Field>
          {isJeugd && <Field label="Groep (basisschool)"><input className={inputCls} style={inputStyle} placeholder="bv. groep 6" value={f.groep} onChange={e => upd('groep', e.target.value)} /></Field>}
        </div>
        <Field label="Inschrijfdatum"><input type="date" className={inputCls} style={inputStyle} value={f.datumInschrijving} onChange={e => upd('datumInschrijving', e.target.value)} /></Field>
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={submit} disabled={!f.naam}>{item ? 'Opslaan' : 'Toevoegen'}</Btn>
      </div>
    </Modal>
  );
}

function BetaalWorkshopForm({ inschrijving, accounts, onSave, onClose }) {
  const [f, setF] = useState({ bedrag: '', datum: new Date().toISOString().slice(0, 10), boeken: true, rekening: '908', grootboekCode: '' });
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  return (
    <Modal title={`Betaling — ${inschrijving.naam}`} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Bedrag"><input type="number" step="0.01" className={inputCls} style={inputStyle} value={f.bedrag} onChange={e => upd('bedrag', e.target.value)} /></Field>
        <Field label="Datum ontvangen"><input type="date" className={inputCls} style={inputStyle} value={f.datum} onChange={e => upd('datum', e.target.value)} /></Field>
        <label className="flex items-start gap-2 text-sm rounded-lg p-2.5" style={{ background: C.paperDim }}>
          <input type="checkbox" checked={f.boeken} onChange={e => upd('boeken', e.target.checked)} className="mt-0.5" />
          <span>Ook automatisch boeken in Financiën</span>
        </label>
        {f.boeken && (
          <>
            <Field label="Op rekening">
              <select className={inputCls} style={inputStyle} value={f.rekening} onChange={e => upd('rekening', e.target.value)}>
                <option value="908">.908 lopend</option>
                <option value="319">.319 spaar</option>
              </select>
            </Field>
            <Field label="Grootboekrekening">
              <GrootboekPicker accounts={accounts} code={f.grootboekCode} onPick={code => upd('grootboekCode', code)} />
            </Field>
          </>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(f)} disabled={!f.bedrag || (f.boeken && !f.grootboekCode)}>Markeer betaald</Btn>
      </div>
    </Modal>
  );
}

/* =========================================================================
   FINANCIËN
========================================================================= */
function FinancienTab({ tx, setTx, accounts, boekjaren, setBoekjaren, members, contributies, setContributies, workshops, inschrijvingen, setInschrijvingen, readOnly, initialQuery, onTrash, onLog }) {
  const years = Array.from(new Set([...boekjaren, ...tx.map(t => t.jaar)])).sort((a, b) => b - a);
  const huidigJaar = new Date().getFullYear();
  const [rekening, setRekening] = useState('908');
  const [jaar, setJaar] = useState(initialQuery ? 'alle' : (years.includes(huidigJaar) ? huidigJaar : (years[0] || huidigJaar)));
  const [maand, setMaand] = useState('alle');
  const [q, setQ] = useState(initialQuery || '');
  const [filterDatum, setFilterDatum] = useState('');
  const [filterBedrag, setFilterBedrag] = useState('');
  const [sortField, setSortField] = useState('datum');
  const [sortDir, setSortDir] = useState('desc');
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [showNewYear, setShowNewYear] = useState(false);
  const [delId, setDelId] = useState(null);

  function toggleSort(field) {
    if (sortField === field) setSortDir(d => (d === 'asc' ? 'desc' : 'asc'));
    else { setSortField(field); setSortDir('asc'); }
  }
  function sortIcon(field) {
    if (sortField !== field) return '';
    return sortDir === 'asc' ? ' \u25B2' : ' \u25BC';
  }

  function addYear(newYear) {
    if (boekjaren.includes(newYear)) { setJaar(newYear); setShowNewYear(false); return; }
    setBoekjaren([...boekjaren, newYear].sort((a, b) => a - b));
    onLog(`Nieuw boekjaar aangemaakt: ${newYear}`, 'financien');
    setJaar(newYear);
    setShowNewYear(false);
  }

  const filtered = tx.filter(t => {
    if (rekening !== 'alle' && t.rekening !== rekening) return false;
    if (jaar !== 'alle' && t.jaar !== Number(jaar)) return false;
    if (maand !== 'alle' && Number(t.maand) !== Number(maand)) return false;
    if (q) {
      const hay = `${t.omschrijving || ''} ${t.grootboek_naam || ''} ${t.grootboek_code || ''}`.toLowerCase();
      if (!hay.includes(q.toLowerCase())) return false;
    }
    if (filterDatum) {
      const hay = `${fmtDate(t.datum)} ${t.datum || ''}`.toLowerCase();
      if (!hay.includes(filterDatum.toLowerCase())) return false;
    }
    if (filterBedrag) {
      const hay = `${euro(t.bedrag)} ${t.bedrag}`.toLowerCase();
      if (!hay.includes(filterBedrag.toLowerCase())) return false;
    }
    return true;
  });

  const sortedFiltered = [...filtered].sort((a, b) => {
    let cmp;
    if (sortField === 'bedrag') cmp = (a.bedrag || 0) - (b.bedrag || 0);
    else if (sortField === 'rekening') cmp = String(a.rekening || '').localeCompare(String(b.rekening || ''));
    else if (sortField === 'grootboek') cmp = gbLabel(a).localeCompare(gbLabel(b));
    else if (sortField === 'omschrijving') cmp = (a.omschrijving || '').localeCompare(b.omschrijving || '');
    else cmp = (a.datum || '').localeCompare(b.datum || '');
    return sortDir === 'asc' ? cmp : -cmp;
  });

  const inkomsten = filtered.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0);
  const uitgaven = filtered.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0);

  function pasKoppelingToe(gekoppeldType, gekoppeldRef, betaald, bedrag) {
    if (gekoppeldType === 'contributie') {
      setContributies(contributies.map(c => (c.lidId === gekoppeldRef.lidId && String(c.jaar) === String(gekoppeldRef.jaar)) ? { ...c, betaald, ...(betaald ? { bedrag: Math.abs(bedrag) } : {}) } : c));
    } else if (gekoppeldType === 'workshop') {
      setInschrijvingen(inschrijvingen.map(i => i.id === gekoppeldRef.inschrijvingId ? { ...i, betaald, ...(betaald ? { bedrag: Math.abs(bedrag) } : {}) } : i));
    }
  }
  function save(data) {
    const payload = { ...data, bedrag: Number(data.bedrag), jaar: new Date(data.datum).getFullYear(), maand: new Date(data.datum).getMonth() + 1 };
    const vorige = data.id ? tx.find(t => t.id === data.id) : null;
    // Was dit gekoppeld en is de koppeling nu weg/veranderd? Dan de oude koppeling eerst terugzetten naar open.
    if (vorige && vorige.gekoppeldType && (!payload.gekoppeldType || vorige.gekoppeldType !== payload.gekoppeldType || JSON.stringify(vorige.gekoppeldRef) !== JSON.stringify(payload.gekoppeldRef))) {
      pasKoppelingToe(vorige.gekoppeldType, vorige.gekoppeldRef, false);
    }
    if (payload.gekoppeldType) pasKoppelingToe(payload.gekoppeldType, payload.gekoppeldRef, true, payload.bedrag);
    if (data.id) { setTx(tx.map(t => t.id === data.id ? payload : t)); onLog(`Boeking bewerkt: ${payload.omschrijving || payload.grootboek_naam} (${euro(payload.bedrag)})`, 'financien'); }
    else { setTx([...tx, { ...payload, id: uid(tx) }]); onLog(`Boeking toegevoegd: ${payload.omschrijving || payload.grootboek_naam} (${euro(payload.bedrag)})${payload.gekoppeldType ? ' — gekoppeld' : ''}`, 'financien'); }
    setShowForm(false); setEditing(null);
  }
  function remove(id) {
    const t = tx.find(x => x.id === id);
    if (t && t.gekoppeldType) pasKoppelingToe(t.gekoppeldType, t.gekoppeldRef, false);
    setTx(tx.filter(t => t.id !== id));
    if (t) { onTrash('boeking', t); onLog(`Boeking verwijderd: ${t.omschrijving || t.grootboek_naam} (${euro(t.bedrag)})${t.gekoppeldType ? ' — koppeling teruggezet naar open' : ''}`, 'financien'); }
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-2 justify-between">
        <div className="flex flex-wrap items-center gap-2">
          <select value={rekening} onChange={e => setRekening(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Beide rekeningen</option>
            <option value="908">Rabobank .908 (lopend)</option>
            <option value="319">Rabobank .319 (spaar)</option>
          </select>
          <select value={jaar} onChange={e => setJaar(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Alle jaren</option>
            {years.map(y => <option key={y} value={y}>{y}</option>)}
          </select>
          {!readOnly && <button onClick={() => setShowNewYear(true)} className="text-xs underline whitespace-nowrap" style={{ color: C.clay }}>+ nieuw boekjaar</button>}
          <select value={maand} onChange={e => setMaand(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Alle maanden</option>
            {MONTH_NAMES.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
          </select>
          <div className="relative">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: C.inkSoft }} />
            <input value={q} onChange={e => setQ(e.target.value)} placeholder="Zoek omschrijving…" className={`${inputCls} pl-8`} style={{ ...inputStyle, width: 190 }} />
          </div>
        </div>
        {!readOnly && <Btn icon={Plus} onClick={() => { setEditing(null); setShowForm(true); }}>Nieuwe boeking</Btn>}
      </div>
      {showNewYear && <NewYearModal existing={years} onSave={addYear} onClose={() => setShowNewYear(false)} label="boekjaar" />}

      <div className="grid grid-cols-3 gap-3">
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Inkomsten</p><p className="font-semibold" style={{ color: C.sageDeep }}>{euro(inkomsten)}</p></Card>
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Uitgaven</p><p className="font-semibold" style={{ color: C.rose }}>{euro(uitgaven)}</p></Card>
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Saldo selectie</p><p className="font-semibold">{euro(inkomsten + uitgaven)}</p></Card>
      </div>

      {jaar !== 'alle' && (() => {
        const sb = saldoBoekjaar(tx, rekening, Number(jaar));
        return (
          <Card className="p-4">
            <h3 className="font-semibold mb-2 text-sm" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>
              Saldo boekjaar {jaar}{rekening !== 'alle' ? ` · rekening .${rekening}` : ' · beide rekeningen'}
            </h3>
            <div className="grid grid-cols-3 gap-3">
              <div className="text-center">
                <p className="text-xs" style={{ color: C.inkSoft }}>Startsaldo (1 jan {jaar})</p>
                <p className="font-semibold">{euro(sb.startsaldo)}</p>
              </div>
              <div className="text-center">
                <p className="text-xs" style={{ color: C.inkSoft }}>Mutaties dit boekjaar</p>
                <p className="font-semibold" style={{ color: sb.mutaties >= 0 ? C.sageDeep : C.rose }}>{sb.mutaties >= 0 ? '+' : ''}{euro(sb.mutaties)}</p>
              </div>
              <div className="text-center">
                <p className="text-xs" style={{ color: C.inkSoft }}>Eindsaldo (huidig/31 dec {jaar})</p>
                <p className="font-semibold">{euro(sb.eindsaldo)}</p>
              </div>
            </div>
            <p className="text-xs mt-2" style={{ color: C.inkSoft }}>Er wordt geen apart startsaldo per boekjaar opgeslagen — dit wordt automatisch afgeleid uit het beginsaldo en alle boekingen tot dit jaar.</p>
          </Card>
        );
      })()}

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              {[
                { key: 'datum', label: 'Datum' },
                { key: 'rekening', label: 'Rek.' },
                { key: 'grootboek', label: 'Grootboek' },
                { key: 'omschrijving', label: 'Omschrijving' },
                { key: 'bedrag', label: 'Bedrag' },
              ].map(col => (
                <th key={col.key} onClick={() => toggleSort(col.key)}
                  className="px-3 py-2 font-medium text-xs cursor-pointer select-none whitespace-nowrap" style={{ color: sortField === col.key ? C.clay : C.inkSoft }}>
                  {col.label}{sortIcon(col.key)}
                </th>
              ))}
              <th></th>
            </tr>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th className="px-3 py-1.5"><input value={filterDatum} onChange={e => setFilterDatum(e.target.value)} placeholder="filter…" className="w-full rounded border px-1.5 py-0.5 text-xs" style={inputStyle} /></th>
              <th className="px-3 py-1.5"></th>
              <th className="px-3 py-1.5" colSpan={2}><input value={q} onChange={e => setQ(e.target.value)} placeholder="filter op grootboek/omschrijving…" className="w-full rounded border px-1.5 py-0.5 text-xs" style={inputStyle} /></th>
              <th className="px-3 py-1.5"><input value={filterBedrag} onChange={e => setFilterBedrag(e.target.value)} placeholder="filter…" className="w-full rounded border px-1.5 py-0.5 text-xs" style={inputStyle} /></th>
              <th className="px-3 py-1.5"></th>
            </tr>
          </thead>
          <tbody>
            {sortedFiltered.slice(0, 300).map(t => (
              <tr key={t.id} className="border-b last:border-0 hover:bg-black/[0.02]" style={{ borderColor: C.border }}>
                <td className="px-3 py-2 whitespace-nowrap" style={{ color: C.inkSoft }}>{fmtDate(t.datum)}</td>
                <td className="px-3 py-2"><Badge tone="muted">.{t.rekening}</Badge></td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>{gbLabel(t)}</td>
                <td className="px-3 py-2">
                  {t.omschrijving}
                  {t.gekoppeldType && <Link2 size={12} className="inline ml-1.5 align-text-top" style={{ color: C.sageDeep }} />}
                </td>
                <td className="px-3 py-2 font-medium whitespace-nowrap" style={{ color: t.bedrag >= 0 ? C.sageDeep : C.rose }}>
                  {t.bedrag >= 0 ? '+' : ''}{euro(t.bedrag)}
                </td>
                <td className="px-3 py-2">
                  <div className="flex gap-1 justify-end">
                    {!readOnly && <>
                      <button onClick={() => { setEditing(t); setShowForm(true); }} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={14} /></button>
                      <button onClick={() => setDelId(t.id)} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={14} /></button>
                    </>}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!filtered.length && <EmptyState icon={Wallet} text="Geen boekingen gevonden met deze filters." />}
        {filtered.length > 300 && <p className="text-xs text-center py-2" style={{ color: C.inkSoft }}>Toont eerste 300 van {filtered.length} — filter verder om te verfijnen.</p>}
      </Card>

      {showForm && <TxForm item={editing} accounts={accounts} members={members} contributies={contributies} workshops={workshops} inschrijvingen={inschrijvingen} onSave={save} onClose={() => { setShowForm(false); setEditing(null); }} />}
      {delId != null && (
        <ConfirmModal message={tx.find(t => t.id === delId)?.gekoppeldType ? 'Deze boeking verwijderen? Dit zet de gekoppelde contributie/inschrijving ook weer terug naar open.' : 'Deze boeking verwijderen?'} onConfirm={() => { remove(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
    </div>
  );
}

function NewYearModal({ existing, onSave, onClose, label = 'boekjaar', copyFrom = null }) {
  const suggestion = existing.length ? Math.max(...existing) + 1 : new Date().getFullYear();
  const [jaar, setJaar] = useState(suggestion);
  const [kopieerVan, setKopieerVan] = useState(copyFrom && copyFrom.length ? copyFrom[copyFrom.length - 1] : '');
  const bestaatAl = existing.includes(Number(jaar));
  return (
    <Modal title={`Nieuw ${label}`} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Jaartal">
          <input type="number" className={inputCls} style={inputStyle} value={jaar} onChange={e => setJaar(e.target.value)} />
        </Field>
        {bestaatAl && <p className="text-xs" style={{ color: C.rose }}>Dit {label} bestaat al — kies een ander jaartal.</p>}
        {copyFrom && (
          <Field label="Beginnen met categorieën van">
            <select className={inputCls} style={inputStyle} value={kopieerVan} onChange={e => setKopieerVan(e.target.value)}>
              <option value="">— leeg beginnen —</option>
              {copyFrom.map(y => <option key={y} value={y}>{y} (bedragen worden op 0 gezet)</option>)}
            </select>
          </Field>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(Number(jaar), kopieerVan || null)} disabled={!jaar || bestaatAl}>Aanmaken</Btn>
      </div>
    </Modal>
  );
}

function GrootboekPicker({ accounts, code, onPick, disabled }) {
  const gekozen = accounts.find(a => String(a.code) === String(code));
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState('');

  const term = q.trim().toLowerCase();
  const resultaten = term
    ? accounts.filter(a => a.code.toLowerCase().includes(term) || a.naam.toLowerCase().includes(term)).slice(0, 30)
    : accounts;

  function kies(a) {
    onPick(a.code);
    setQ('');
    setOpen(false);
  }

  return (
    <div className="relative">
      <input
        disabled={disabled}
        className={inputCls}
        style={inputStyle}
        placeholder="Zoek op nummer of omschrijving…"
        value={open ? q : (gekozen ? `${gekozen.code}  ${gekozen.naam}` : '')}
        onFocus={() => { setOpen(true); setQ(''); }}
        onChange={e => setQ(e.target.value)}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
      />
      {open && !disabled && (
        <div className="absolute left-0 right-0 mt-1 max-h-56 overflow-y-auto rounded-lg border shadow-lg z-50" style={{ background: C.card, borderColor: C.border }}>
          {resultaten.map(a => (
            <button key={a.code} type="button" onMouseDown={() => kies(a)}
              className="w-full text-left text-sm px-3 py-1.5 hover:bg-black/5">
              <span style={{ color: C.inkSoft }}>{a.code}</span>  {a.naam}
            </button>
          ))}
          {!resultaten.length && <p className="text-xs px-3 py-2" style={{ color: C.inkSoft }}>Geen grootboekrekening gevonden.</p>}
        </div>
      )}
    </div>
  );
}

/* Bouwt de lijst openstaande contributies/inschrijvingen voor de koppelsectie bij een
   nieuwe boeking in Financiën, gesorteerd op beste match (bedrag + naam in de omschrijving). */
function bouwKoppelKandidaten(contributies, members, inschrijvingen, workshops, bedrag, omschrijving) {
  const tekst = (omschrijving || '').toLowerCase();
  const doelBedrag = Math.abs(Number(bedrag) || 0);
  const lijst = [];
  contributies.filter(c => !c.betaald).forEach(c => {
    const lid = members.find(m => m.id === c.lidId);
    if (!lid) return;
    lijst.push({ key: `contributie-${c.id}`, type: 'contributie', ref: { lidId: c.lidId, jaar: c.jaar }, naam: fullName(lid), label: `Contributie ${c.jaar} — ${fullName(lid)}`, bedrag: c.bedrag });
  });
  inschrijvingen.filter(i => !i.betaald).forEach(i => {
    const ws = workshops.find(w => w.id === i.workshopId);
    lijst.push({ key: `workshop-${i.id}`, type: 'workshop', ref: { inschrijvingId: i.id }, naam: i.naam, label: `${ws ? ws.titel : 'Workshop'} — ${i.naam}`, bedrag: i.bedrag });
  });
  return lijst.map(k => {
    let score = 0;
    if (k.bedrag != null && Math.abs(Number(k.bedrag)) === doelBedrag && doelBedrag > 0) score += 2;
    if (tekst && k.naam && tekst.includes(k.naam.toLowerCase())) score += 1;
    return { ...k, score };
  }).sort((a, b) => b.score - a.score || a.label.localeCompare(b.label));
}

function TxForm({ item, accounts, members, contributies, workshops, inschrijvingen, onSave, onClose }) {
  const [f, setF] = useState(() => item ? { ...item } : {
    id: null, rekening: '908', datum: new Date().toISOString().slice(0, 10),
    grootboek_code: '', grootboek_naam: '', bedrag: '', omschrijving: '',
    gekoppeldType: null, gekoppeldRef: null,
  });
  const [toonKoppelen, setToonKoppelen] = useState(!!f.gekoppeldType);
  const [koppelZoek, setKoppelZoek] = useState('');
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  function pickGrootboek(code) {
    const acc = accounts.find(a => String(a.code) === code);
    setF(s => ({ ...s, grootboek_code: code, grootboek_naam: acc ? acc.naam : s.grootboek_naam }));
  }
  const kandidaten = bouwKoppelKandidaten(contributies, members, inschrijvingen, workshops, f.bedrag, f.omschrijving)
    .filter(k => !koppelZoek.trim() || k.label.toLowerCase().includes(koppelZoek.trim().toLowerCase()));
  const gekoppeldeNaam = (() => {
    if (!f.gekoppeldType) return null;
    if (f.gekoppeldType === 'contributie') {
      const lid = members.find(m => m.id === f.gekoppeldRef.lidId);
      return lid ? `Contributie ${f.gekoppeldRef.jaar} — ${fullName(lid)}` : null;
    }
    const i = inschrijvingen.find(x => x.id === f.gekoppeldRef.inschrijvingId);
    return i ? `Workshopinschrijving — ${i.naam}` : null;
  })();

  return (
    <Modal title={item ? 'Boeking bewerken' : 'Nieuwe boeking'} onClose={onClose}>
      <div className="space-y-3">
        <div className="grid grid-cols-2 gap-3">
          <Field label="Rekening">
            <select className={inputCls} style={inputStyle} value={f.rekening} onChange={e => upd('rekening', e.target.value)}>
              <option value="908">.908 lopend</option>
              <option value="319">.319 spaar</option>
            </select>
          </Field>
          <Field label="Datum"><input type="date" className={inputCls} style={inputStyle} value={f.datum || ''} onChange={e => upd('datum', e.target.value)} /></Field>
        </div>
        <Field label="Grootboekrekening">
          <GrootboekPicker accounts={accounts} code={f.grootboek_code} onPick={pickGrootboek} />
        </Field>
        <Field label="Bedrag (negatief = uitgave)"><input type="number" step="0.01" className={inputCls} style={inputStyle} value={f.bedrag ?? ''} onChange={e => upd('bedrag', e.target.value)} /></Field>
        <Field label="Omschrijving"><input className={inputCls} style={inputStyle} value={f.omschrijving || ''} onChange={e => upd('omschrijving', e.target.value)} /></Field>

        {!toonKoppelen && !f.gekoppeldType && (
          <button type="button" onClick={() => { setToonKoppelen(true); setKoppelZoek(f.omschrijving || ''); }} className="text-xs underline" style={{ color: C.clay }}>+ Koppelen aan openstaande contributie of workshopinschrijving</button>
        )}
        {f.gekoppeldType && (
          <div className="flex items-center justify-between text-sm rounded-lg px-3 py-2" style={{ background: C.paperDim }}>
            <span>Gekoppeld aan: <strong>{gekoppeldeNaam}</strong></span>
            <button type="button" onClick={() => upd('gekoppeldType', null)} className="text-xs underline" style={{ color: C.rose }}>loskoppelen</button>
          </div>
        )}
        {toonKoppelen && !f.gekoppeldType && (
          <div className="space-y-2">
            <input className={inputCls} style={inputStyle} placeholder="Zoek op naam…" value={koppelZoek} onChange={e => setKoppelZoek(e.target.value)} />
            <div className="max-h-48 overflow-y-auto space-y-1">
              {kandidaten.slice(0, 30).map(k => (
                <button type="button" key={k.key} onClick={() => setF(s => ({ ...s, gekoppeldType: k.type, gekoppeldRef: k.ref }))}
                  className="w-full text-left text-sm px-2.5 py-1.5 rounded-lg hover:bg-black/5 flex items-center justify-between"
                  style={{ background: k.score > 0 ? C.paperDim : 'transparent' }}>
                  <span>{k.label}</span>
                  {k.bedrag != null && <span style={{ color: C.inkSoft }}>{euro(k.bedrag)}</span>}
                </button>
              ))}
              {!kandidaten.length && <p className="text-xs italic px-2" style={{ color: C.inkSoft }}>Geen openstaande contributies/inschrijvingen gevonden.</p>}
            </div>
            <button type="button" onClick={() => setToonKoppelen(false)} className="text-xs underline" style={{ color: C.inkSoft }}>niet koppelen</button>
          </div>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(f)} disabled={!f.grootboek_code || !f.bedrag}>Opslaan</Btn>
      </div>
    </Modal>
  );
}

/* =========================================================================
   BEGROTING
========================================================================= */
function BegrotingTab({ budget, setBudget, tx, boekjaren, setBoekjaren, begrotingKoppelingen, readOnly, onLog }) {
  const jarenLijst = Array.from(new Set([...boekjaren, ...Object.keys(budget).map(Number)])).sort((a, b) => a - b).map(String);
  const huidigJaar = String(new Date().getFullYear());
  const [jaar, setJaar] = useState(jarenLijst.includes(huidigJaar) ? huidigJaar : (jarenLijst[0] || huidigJaar));
  const [showForm, setShowForm] = useState(false);
  const [showNewYear, setShowNewYear] = useState(false);
  const [delIdx, setDelIdx] = useState(null);
  const [wisType, setWisType] = useState(null);
  const rows = budget[jaar] || [];

  function actualFor(categorie) {
    const codes = budgetGbCodes(categorie, begrotingKoppelingen);
    if (!codes.length) return null;
    const sum = tx.filter(t => codes.includes(t.grootboek_code) && String(t.jaar) === jaar).reduce((s, t) => s + t.bedrag, 0);
    return sum;
  }

  function updateRow(idx, patch) {
    const next = [...rows]; next[idx] = { ...next[idx], ...patch };
    setBudget({ ...budget, [jaar]: next });
  }
  function removeRow(idx) {
    onLog(`Begrotingsregel verwijderd: ${rows[idx].categorie} (${jaar})`, 'begroting');
    setBudget({ ...budget, [jaar]: rows.filter((_, i) => i !== idx) });
  }
  function addRow(row) {
    setBudget({ ...budget, [jaar]: [...rows, row] });
    onLog(`Begrotingsregel toegevoegd: ${row.categorie} (${jaar})`, 'begroting');
    setShowForm(false);
  }
  function addYear(newYear, kopieerVan) {
    const key = String(newYear);
    if (budget[key]) { setJaar(key); setShowNewYear(false); return; }
    const bron = kopieerVan && budget[kopieerVan] ? budget[kopieerVan] : [];
    const nieuweRegels = bron.map(r => ({ ...r, bedrag: 0 }));
    setBudget({ ...budget, [key]: nieuweRegels });
    if (!boekjaren.includes(newYear)) setBoekjaren([...boekjaren, newYear].sort((a, b) => a - b));
    onLog(`Nieuw begrotingsjaar aangemaakt: ${newYear}`, 'begroting');
    setJaar(key);
    setShowNewYear(false);
  }

  const heeftBoekingen = jarenLijst.reduce((acc, j) => {
    acc[j] = tx.some(t => String(t.jaar) === j);
    return acc;
  }, {});
  function magWissen(j) {
    return j !== huidigJaar && !heeftBoekingen[j];
  }
  function reden(j) {
    if (j === huidigJaar) return 'Het lopende boekjaar kan niet gewist worden.';
    if (heeftBoekingen[j]) return 'Er staan al boekingen op dit jaar in Financiën — eerst die verwijderen.';
    return '';
  }
  function wisBegroting(j) {
    const rest = { ...budget };
    delete rest[j];
    setBudget(rest);
    onLog(`Begroting gewist: boekjaar ${j}`, 'begroting');
    const overigeJaren = jarenLijst.filter(x => x !== j);
    setJaar(overigeJaren[0] || huidigJaar);
  }
  function wisBoekjaar(j) {
    const rest = { ...budget };
    delete rest[j];
    setBudget(rest);
    setBoekjaren(boekjaren.filter(y => String(y) !== j));
    onLog(`Boekjaar volledig verwijderd: ${j}`, 'begroting');
    const overigeJaren = jarenLijst.filter(x => x !== j);
    setJaar(overigeJaren[0] || huidigJaar);
  }

  const inkomstenRows = rows.map((r, i) => ({ ...r, i })).filter(r => r.sectie === 'Inkomsten');
  const uitgavenRows = rows.map((r, i) => ({ ...r, i })).filter(r => r.sectie === 'Uitgaven');
  const totInk = inkomstenRows.reduce((s, r) => s + Number(r.bedrag || 0), 0);
  const totUit = uitgavenRows.reduce((s, r) => s + Number(r.bedrag || 0), 0);

  function Section({ title, list, tone }) {
    return (
      <Card className="p-4">
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-semibold" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>{title}</h3>
          <Badge tone={tone}>{euro(list.reduce((s, r) => s + Number(r.bedrag || 0), 0))}</Badge>
        </div>
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th className="py-1.5 font-medium text-xs" style={{ color: C.inkSoft }}>Categorie</th>
              <th className="py-1.5 font-medium text-xs text-right" style={{ color: C.inkSoft }}>Begroot</th>
              <th className="py-1.5 font-medium text-xs text-right" style={{ color: C.inkSoft }}>Werkelijk</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {list.map(r => {
              const actual = actualFor(r.categorie);
              return (
                <tr key={r.i} className="border-b last:border-0" style={{ borderColor: C.border }}>
                  <td className="py-1.5 pr-2">{r.categorie}{r.notitie && <span className="block text-xs" style={{ color: C.inkSoft }}>{r.notitie}</span>}</td>
                  <td className="py-1.5 text-right">
                    <DebouncedField type="number" value={r.bedrag ?? 0} disabled={readOnly}
                      onCommit={v => updateRow(r.i, { bedrag: v === '' ? '' : Number(v) })}
                      className="w-24 text-right rounded border px-1.5 py-0.5 text-sm" style={inputStyle} />
                  </td>
                  <td className="py-1.5 text-right" style={{ color: actual == null ? C.inkSoft : (actual >= 0 ? C.sageDeep : C.rose) }}>
                    {actual == null ? '—' : euro(actual)}
                  </td>
                  <td className="py-1.5 pl-2 text-right">{!readOnly && <button onClick={() => setDelIdx(r.i)} style={{ color: C.rose }}><Trash2 size={13} /></button>}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {(() => {
          const metCode = list.filter(r => actualFor(r.categorie) != null);
          const zonderCode = list.length - metCode.length;
          const totaalWerkelijk = metCode.reduce((s, r) => s + actualFor(r.categorie), 0);
          return (
            <div className="flex items-center justify-between mt-2 pt-2 border-t text-sm" style={{ borderColor: C.border }}>
              <span style={{ color: C.inkSoft }}>Totaal werkelijk{zonderCode > 0 ? ` (${zonderCode} categorie${zonderCode > 1 ? 'ën' : ''} niet gekoppeld, niet meegeteld)` : ''}</span>
              <span className="font-semibold" style={{ color: totaalWerkelijk >= 0 ? C.sageDeep : C.rose }}>{euro(totaalWerkelijk)}</span>
            </div>
          );
        })()}
      </Card>
    );
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="flex items-center gap-1 flex-wrap">
          {jarenLijst.map(j => (
            <button key={j} onClick={() => setJaar(j)} className="px-3 py-1.5 rounded-lg text-sm font-medium border"
              style={{ background: jaar === j ? C.clay : 'transparent', color: jaar === j ? '#fff' : C.ink, borderColor: jaar === j ? C.clay : C.border }}>
              {j}
            </button>
          ))}
          {!readOnly && <button onClick={() => setShowNewYear(true)} className="text-xs underline ml-1" style={{ color: C.clay }}>+ nieuw boekjaar / begroting</button>}
        </div>
        {!readOnly && <Btn icon={Plus} onClick={() => setShowForm(true)}>Nieuwe begrotingsregel</Btn>}
      </div>

      {!readOnly && (
        magWissen(jaar) ? (
          <div className="flex flex-wrap gap-2">
            <button onClick={() => setWisType('begroting')} className="text-xs underline" style={{ color: C.rose }}>Begroting {jaar} wissen</button>
            <button onClick={() => setWisType('boekjaar')} className="text-xs underline" style={{ color: C.rose }}>Boekjaar {jaar} volledig verwijderen</button>
          </div>
        ) : (
          <p className="text-xs" style={{ color: C.inkSoft }}>{reden(jaar)}</p>
        )
      )}

      <div className="grid sm:grid-cols-3 gap-3">
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Begrote inkomsten</p><p className="font-semibold" style={{ color: C.sageDeep }}>{euro(totInk)}</p></Card>
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Begrote uitgaven</p><p className="font-semibold" style={{ color: C.rose }}>{euro(totUit)}</p></Card>
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Begroot resultaat</p><p className="font-semibold">{euro(totInk - totUit)}</p></Card>
      </div>

      {(() => {
        const sb = saldoBoekjaar(tx, 'alle', Number(jaar));
        return (
          <Card className="p-4">
            <h3 className="font-semibold mb-2 text-sm" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Werkelijk saldo boekjaar {jaar} (beide rekeningen)</h3>
            <div className="grid grid-cols-3 gap-3">
              <div className="text-center">
                <p className="text-xs" style={{ color: C.inkSoft }}>Startsaldo (1 jan {jaar})</p>
                <p className="font-semibold">{euro(sb.startsaldo)}</p>
              </div>
              <div className="text-center">
                <p className="text-xs" style={{ color: C.inkSoft }}>Mutaties dit boekjaar</p>
                <p className="font-semibold" style={{ color: sb.mutaties >= 0 ? C.sageDeep : C.rose }}>{sb.mutaties >= 0 ? '+' : ''}{euro(sb.mutaties)}</p>
              </div>
              <div className="text-center">
                <p className="text-xs" style={{ color: C.inkSoft }}>Eindsaldo (huidig/31 dec {jaar})</p>
                <p className="font-semibold">{euro(sb.eindsaldo)}</p>
              </div>
            </div>
            <p className="text-xs mt-2" style={{ color: C.inkSoft }}>Gebaseerd op de werkelijke boekingen in Financiën — dit is los van de begrote bedragen hierboven.</p>
          </Card>
        );
      })()}

      <div className="grid lg:grid-cols-2 gap-4">
        <Section title="Inkomsten" list={inkomstenRows} tone="sage" />
        <Section title="Uitgaven" list={uitgavenRows} tone="clay" />
      </div>

      {showForm && (
        <Modal title="Begrotingsregel toevoegen" onClose={() => setShowForm(false)}>
          <BudgetRowForm onSave={addRow} onClose={() => setShowForm(false)} />
        </Modal>
      )}
      {showNewYear && (
        <NewYearModal existing={jarenLijst.map(Number)} onSave={addYear} onClose={() => setShowNewYear(false)}
          label="boekjaar" copyFrom={Object.keys(budget).filter(y => budget[y] && budget[y].length)} />
      )}
      {delIdx != null && (
        <ConfirmModal message="Deze begrotingsregel verwijderen?" onConfirm={() => { removeRow(delIdx); setDelIdx(null); }} onCancel={() => setDelIdx(null)} />
      )}
      {wisType === 'begroting' && (
        <ConfirmModal title="Begroting wissen" message={`Weet je zeker dat je de volledige begroting van ${jaar} wilt wissen? Het boekjaar zelf blijft bestaan.`}
          onConfirm={() => { wisBegroting(jaar); setWisType(null); }} onCancel={() => setWisType(null)} />
      )}
      {wisType === 'boekjaar' && (
        <ConfirmModal title="Boekjaar verwijderen" message={`Weet je zeker dat je boekjaar ${jaar} (inclusief de begroting) volledig wilt verwijderen? Dit kan niet ongedaan gemaakt worden.`}
          onConfirm={() => { wisBoekjaar(jaar); setWisType(null); }} onCancel={() => setWisType(null)} />
      )}
    </div>
  );
}

function BudgetRowForm({ onSave, onClose }) {
  const [f, setF] = useState({ sectie: 'Uitgaven', categorie: '', bedrag: '', notitie: '' });
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  return (
    <div className="space-y-3">
      <Field label="Sectie">
        <select className={inputCls} style={inputStyle} value={f.sectie} onChange={e => upd('sectie', e.target.value)}>
          <option value="Inkomsten">Inkomsten</option>
          <option value="Uitgaven">Uitgaven</option>
        </select>
      </Field>
      <Field label="Categorie"><input className={inputCls} style={inputStyle} placeholder="bv. 4300  Huur" value={f.categorie} onChange={e => upd('categorie', e.target.value)} /></Field>
      <Field label="Begroot bedrag"><input type="number" step="0.01" className={inputCls} style={inputStyle} value={f.bedrag} onChange={e => upd('bedrag', e.target.value)} /></Field>
      <Field label="Notitie (optioneel)"><input className={inputCls} style={inputStyle} value={f.notitie} onChange={e => upd('notitie', e.target.value)} /></Field>
      <div className="flex justify-end gap-2 pt-2">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave({ ...f, bedrag: Number(f.bedrag) || 0 })} disabled={!f.categorie}>Toevoegen</Btn>
      </div>
    </div>
  );
}

/* =========================================================================
   RAPPORTAGE
========================================================================= */
function RapportageTab({ tx, accounts, members, workshops, inschrijvingen, budget, begrotingKoppelingen, logoHoogteCm }) {
  const years = Array.from(new Set(tx.map(t => t.jaar))).sort((a, b) => b - a);
  const huidigJaar = new Date().getFullYear();
  const [jaar, setJaar] = useState(years.includes(huidigJaar) ? huidigJaar : (years[0] || huidigJaar));
  const [maand, setMaand] = useState('alle');

  const scope = tx.filter(t => t.jaar === Number(jaar) && (maand === 'alle' || Number(t.maand) === Number(maand)));

  const perGrootboek = useMemo(() => {
    const map = {};
    scope.forEach(t => {
      const key = gbLabel(t) || 'Onbekend';
      map[key] = (map[key] || 0) + t.bedrag;
    });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [scope]);

  const inkomsten = scope.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0);
  const uitgaven = scope.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0);

  const perMaand = MONTH_NAMES.map((naam, i) => {
    const rows = tx.filter(t => t.jaar === Number(jaar) && Number(t.maand) === i + 1);
    return {
      naam,
      Inkomsten: Math.round(rows.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0)),
      Uitgaven: Math.round(-rows.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0)),
    };
  });

  function titelVoorPeriode() {
    return maand === 'alle' ? `Jaarverslag ${jaar}` : `Maandverslag ${MONTH_NAMES[maand - 1]} ${jaar}`;
  }

  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const titel = titelVoorPeriode();
    const summaryAoA = [
      [titel],
      [],
      ['Rijlabel', 'Bedrag'],
      ...perGrootboek.map(([k, v]) => [k, Number(v.toFixed(2))]),
      [],
      ['Inkomsten', Number(inkomsten.toFixed(2))],
      ['Uitgaven', Number(uitgaven.toFixed(2))],
      ['Resultaat', Number((inkomsten + uitgaven).toFixed(2))],
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(summaryAoA);
    XLSX.utils.book_append_sheet(wb, ws1, 'Samenvatting');

    const txAoA = [
      ['Datum', 'Rekening', 'Grootboek', 'Omschrijving', 'Bedrag'],
      ...scope.sort((a, b) => (a.datum || '').localeCompare(b.datum || '')).map(t => [
        t.datum, '.' + t.rekening, gbLabel(t), t.omschrijving || '', Number(Number(t.bedrag).toFixed(2))
      ]),
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(txAoA);
    XLSX.utils.book_append_sheet(wb, ws2, 'Boekingen');

    XLSX.writeFile(wb, `BladelsCreatief_${titel.replace(/\s+/g, '_')}.xlsx`);
  }

  function exportPdf() {
    const titel = titelVoorPeriode();
    const subtitel = `Inkomsten: ${euro(inkomsten)} · Uitgaven: ${euro(uitgaven)} · Resultaat: ${euro(inkomsten + uitgaven)}`;
    const header = ['Grootboekrekening', 'Bedrag'];
    const rows = perGrootboek.map(([k, v]) => [k, euro(v)]);
    downloadPrintableHtml({ titel, subtitel, header, rows });
  }

  function exportWord() {
    const titel = titelVoorPeriode();
    const gbRows = perGrootboek.map(([k, v]) => `<tr><td>${escapeHtml(k)}</td><td>${euro(v)}</td></tr>`).join('');
    const txRows = scope.sort((a, b) => (a.datum || '').localeCompare(b.datum || '')).map(t =>
      `<tr><td>${fmtDate(t.datum)}</td><td>.${t.rekening}</td><td>${escapeHtml(gbLabel(t))}</td><td>${escapeHtml(t.omschrijving || '')}</td><td>${euro(t.bedrag)}</td></tr>`
    ).join('');
    const body = `${logoImgTag(logoHoogteCm)}<div class="bar"></div>
<h1>${escapeHtml(titel)}</h1>
<p class="meta">Inkomsten: ${euro(inkomsten)} · Uitgaven: ${euro(uitgaven)} · Resultaat: ${euro(inkomsten + uitgaven)}</p>
<h2>Resultaat per grootboekrekening</h2>
<table><tr><th>Grootboekrekening</th><th>Bedrag</th></tr>${gbRows || '<tr><td colspan="2"><em>Geen boekingen</em></td></tr>'}</table>
<h2>Overzicht boekingen</h2>
<table><tr><th>Datum</th><th>Rekening</th><th>Grootboek</th><th>Omschrijving</th><th>Bedrag</th></tr>${txRows || '<tr><td colspan="5"><em>Geen boekingen</em></td></tr>'}</table>`;
    downloadWordDoc({ titel, filename: `BladelsCreatief_${titel.replace(/\s+/g, '_')}.doc`, bodyHtml: body });
  }

  function exportJaarverslag() {
    const jaarNum = Number(jaar);
    const actieveLeden = members.filter(m => m.status === 'actief');
    const workshopDeelnemers = new Set(inschrijvingen.filter(i => i.status !== 'geannuleerd').map(i => i.naam)).size;
    const jaarTx = tx.filter(t => t.jaar === jaarNum);
    const jInkomsten = jaarTx.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0);
    const jUitgaven = jaarTx.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0);
    const map = {};
    jaarTx.forEach(t => { const k = gbLabel(t) || 'Onbekend'; map[k] = (map[k] || 0) + t.bedrag; });
    const gbRows = Object.entries(map).sort((a, b) => b[1] - a[1]).map(([k, v]) => `<tr><td>${escapeHtml(k)}</td><td>${euro(v)}</td></tr>`).join('');

    const bJaar = budget[String(jaarNum)] || [];
    const begrotingRows = bJaar.map(r => {
      const codes = budgetGbCodes(r.categorie, begrotingKoppelingen);
      const werkelijk = codes.length ? jaarTx.filter(t => codes.includes(t.grootboek_code)).reduce((s, t) => s + t.bedrag, 0) : null;
      return `<tr><td>${escapeHtml(r.sectie)}</td><td>${escapeHtml(r.categorie)}</td><td>${euro(r.bedrag)}</td><td>${werkelijk == null ? '—' : euro(werkelijk)}</td></tr>`;
    }).join('');

    const body = `${logoImgTag(logoHoogteCm)}<div class="bar"></div>
<h1>Jaarverslag ${jaarNum}</h1>
<p class="meta">BladelsCreatief · gegenereerd op ${new Date().toLocaleDateString('nl-NL')}</p>
<h2>Leden</h2>
<p>Actieve leden: ${actieveLeden.length} · Inactieve leden: ${members.length - actieveLeden.length} · Unieke workshopdeelnemers: ${workshopDeelnemers}</p>
<h2>Financieel overzicht ${jaarNum}</h2>
<p>Inkomsten: ${euro(jInkomsten)} · Uitgaven: ${euro(jUitgaven)} · Resultaat: ${euro(jInkomsten + jUitgaven)}</p>
<table><tr><th>Grootboekrekening</th><th>Bedrag</th></tr>${gbRows || '<tr><td colspan="2"><em>Geen boekingen</em></td></tr>'}</table>
<h2>Begroting versus werkelijk ${jaarNum}</h2>
${bJaar.length
  ? `<table><tr><th>Sectie</th><th>Categorie</th><th>Begroot</th><th>Werkelijk</th></tr>${begrotingRows}</table>`
  : `<p><em>Geen begroting vastgelegd voor ${jaarNum}.</em></p>`}`;

    downloadWordDoc({ titel: `Jaarverslag ${jaarNum}`, filename: `BladelsCreatief_Jaarverslag_${jaarNum}.doc`, bodyHtml: body });
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-2 justify-between">
        <div className="flex flex-wrap items-center gap-2">
          <select value={jaar} onChange={e => setJaar(e.target.value)} className={inputCls} style={inputStyle}>
            {years.map(y => <option key={y} value={y}>{y}</option>)}
          </select>
          <select value={maand} onChange={e => setMaand(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Heel jaar</option>
            {MONTH_NAMES.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
          </select>
        </div>
        <div className="flex gap-2 flex-wrap">
          <Btn icon={Printer} tone="outline" onClick={exportPdf}>Als PDF</Btn>
          <Btn icon={FileText} tone="outline" onClick={exportWord}>Als Word</Btn>
          <Btn icon={Download} tone="sage" onClick={exportExcel}>Als Excel</Btn>
          <Btn icon={FileText} tone="clay" onClick={exportJaarverslag}>Compleet jaarverslag (Word)</Btn>
        </div>
      </div>

      <div className="grid sm:grid-cols-3 gap-3">
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Inkomsten</p><p className="font-semibold" style={{ color: C.sageDeep }}>{euro(inkomsten)}</p></Card>
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Uitgaven</p><p className="font-semibold" style={{ color: C.rose }}>{euro(uitgaven)}</p></Card>
        <Card className="p-3 text-center"><p className="text-xs" style={{ color: C.inkSoft }}>Resultaat</p><p className="font-semibold">{euro(inkomsten + uitgaven)}</p></Card>
      </div>

      <Card className="p-4">
        <h3 className="font-semibold mb-3" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Maandoverzicht {jaar}</h3>
        <div style={{ height: 240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={perMaand}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} vertical={false} />
              <XAxis dataKey="naam" tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.inkSoft }} axisLine={false} tickLine={false} width={40} />
              <Tooltip formatter={(v) => euro(v)} contentStyle={{ borderRadius: 8, borderColor: C.border, fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Line type="monotone" dataKey="Inkomsten" stroke={C.sage} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="Uitgaven" stroke={C.clay} strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </Card>

      <Card className="p-4">
        <h3 className="font-semibold mb-3" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Resultaat per grootboekrekening</h3>
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th className="py-1.5 font-medium text-xs" style={{ color: C.inkSoft }}>Grootboekrekening</th>
              <th className="py-1.5 font-medium text-xs text-right" style={{ color: C.inkSoft }}>Bedrag</th>
            </tr>
          </thead>
          <tbody>
            {perGrootboek.map(([k, v]) => (
              <tr key={k} className="border-b last:border-0" style={{ borderColor: C.border }}>
                <td className="py-1.5">{k}</td>
                <td className="py-1.5 text-right font-medium" style={{ color: v >= 0 ? C.sageDeep : C.rose }}>{euro(v)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {!perGrootboek.length && <EmptyState icon={FileSpreadsheet} text="Geen boekingen in deze periode." />}
      </Card>
    </div>
  );
}

/* =========================================================================
   INSTELLINGEN — rolbeheer, logboek, prullenbak, back-up
========================================================================= */
function InstellingenTab({ isVoorzitter, ingelogd, rolpermissies, setRolpermissies, beveiliging, setBeveiliging, standaarden, setStandaarden, logboek, prullenbak, setPrullenbak,
  members, setMembers, workshops, setWorkshops, inschrijvingen, setInschrijvingen, tx, setTx, boekjaren, setBoekjaren, accounts, setAccounts,
  begrotingKoppelingen, setBegrotingKoppelingen,
  dagdelen, setDagdelen,
  workshopSoorten, setWorkshopSoorten, agendapuntenVooraf, setAgendapuntenVooraf, agendapuntenAfsluitend, setAgendapuntenAfsluitend,
  budget, setBudget, actielijst, setActielijst, contributies, setContributies, vergaderingen, setVergaderingen,
  magLeden, magWorkshops, magFinancien, magVergaderingen,
  magLedenImporteren, magWorkshopsImporteren, magFinancienImporteren, backupData, onLog }) {
  const [sectie, setSectie] = useState('rollen');
  const [actieveAnderen, setActieveAnderen] = useState(null);

  async function checkActieveGebruikers() {
    try {
      const res = await opslagLezen(STORAGE_KEYS.sessies, true);
      const sessies = res && res.value ? JSON.parse(res.value) : {};
      const grens = Date.now() - 3 * 60 * 1000;
      const eigenNaam = ingelogd ? fullName(ingelogd) : null;
      const anderen = Object.entries(sessies)
        .filter(([naam, laatst]) => naam !== eigenNaam && new Date(laatst).getTime() > grens)
        .map(([naam]) => naam);
      setActieveAnderen(anderen);
    } catch (e) {
      setActieveAnderen([]);
    }
  }
  useEffect(() => {
    if (sectie === 'backup') checkActieveGebruikers();
  }, [sectie]);

  const SECTIES = [
    { id: 'rollen', label: 'Rolbeheer', icon: Settings },
    { id: 'beveiliging', label: 'Beveiliging', icon: Lock },
    { id: 'standaarden', label: 'Standaarden', icon: Sliders },
    { id: 'grootboek', label: 'Grootboekrekeningen', icon: Wallet },
    { id: 'koppelingen', label: 'Begroting-koppelingen', icon: Link2 },
    { id: 'dagdelen', label: 'Dagdelen/Groepen', icon: Grid3x3 },
    { id: 'soorten', label: 'Workshop-soorten', icon: Palette },
    { id: 'agenda', label: 'Standaard agendapunten', icon: ClipboardList },
    { id: 'logboek', label: 'Logboek', icon: History },
    { id: 'prullenbak', label: 'Prullenbak', icon: Trash2 },
    { id: 'importeren', label: 'Importeren', icon: Upload },
    { id: 'backup', label: 'Back-up', icon: DatabaseBackup },
  ];

  function herstel(entry) {
    if (entry.type === 'lid') setMembers([...members, entry.data]);
    if (entry.type === 'workshop') setWorkshops([...workshops, entry.data]);
    if (entry.type === 'inschrijving') setInschrijvingen([...inschrijvingen, entry.data]);
    if (entry.type === 'boeking') setTx([...tx, entry.data]);
    if (entry.type === 'vergadering') setVergaderingen([...vergaderingen, entry.data]);
    setPrullenbak(prullenbak.filter(e => e.id !== entry.id));
    onLog(`Hersteld uit prullenbak: ${entry.type} — ${trashLabel(entry)}`, 'instellingen');
  }
  function verwijderDefinitief(id) {
    setPrullenbak(prullenbak.filter(e => e.id !== id));
  }
  function trashLabel(entry) {
    if (entry.type === 'lid') return fullName(entry.data);
    if (entry.type === 'workshop') return entry.data.titel;
    if (entry.type === 'inschrijving') return entry.data.naam;
    if (entry.type === 'boeking') return entry.data.omschrijving || entry.data.grootboek_naam;
    if (entry.type === 'vergadering') return entry.data.titel;
    return '—';
  }

  function exportBackup() {
    const json = JSON.stringify({ export_datum: new Date().toISOString(), ...backupData }, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `BladelsCreatief_Backup_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    exportBackupExcel();
  }

  function exportBackupExcel() {
    const { members: m, workshops: w, inschrijvingen: i, tx: t, accounts: acc, budget: bud, boekjaren: bj, vergaderingen: verg, actielijst: act, contributies: contr } = backupData;
    const wb = XLSX.utils.book_new();
    const voegToe = (naam, aoa) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), naam.slice(0, 31));

    voegToe('Leden', [
      ['Voornaam', 'Tussenvoegsel', 'Achternaam', 'E-mail', 'Telefoon', 'Adres', 'Postcode', 'Woonplaats', 'Geboortedatum', 'Lid sinds', 'Einde lidmaatschap', 'Status', 'Functie', 'Groepsapp', 'Dagdelen'],
      ...m.map(x => [x.voornaam, x.tussenvoegsel, x.achternaam, x.email, x.telefoon, x.adres, x.postcode, x.woonplaats, x.gebdatum, x.lidsinds, x.eindelidmaat, x.status, x.functie, x.groepsapp, (x.dagdelen || []).join(', ')]),
    ]);

    voegToe('Workshops', [
      ['Titel', 'Soort', 'Type', 'Data', 'Dagdeel', 'Locatie', 'Bedrag', 'Max. deelnemers', 'Status'],
      ...w.map(x => [x.titel, x.soort, x.type, (x.datums || []).join(', '), x.dagdeel, x.locatie, x.bedrag, x.maxDeelnemers, x.status]),
    ]);

    voegToe('Workshop-inschrijvingen', [
      ['Workshop', 'Naam', 'Herkomst', 'E-mail', 'E-mail ouder 2', 'Telefoon', 'Telefoon ouder 2', 'Leeftijd', 'Groep', 'Status', 'Betaald', 'Bedrag', 'Inschrijfdatum'],
      ...i.map(x => {
        const ws = w.find(y => y.id === x.workshopId);
        return [ws ? ws.titel : x.workshopId, x.naam, x.herkomst, x.email, x.email2, x.telefoon, x.telefoon2, x.leeftijd, x.groep, x.status, x.betaald ? 'ja' : 'nee', x.bedrag, x.datumInschrijving];
      }),
    ]);

    voegToe('Financien', [
      ['Rekening', 'Jaar', 'Maand', 'Datum', 'Grootboekcode', 'Grootboeknaam', 'Bedrag', 'Omschrijving'],
      ...t.map(x => ['.' + x.rekening, x.jaar, x.maand, x.datum, x.grootboek_code, x.grootboek_naam, x.bedrag, x.omschrijving]),
    ]);

    voegToe('Grootboekrekeningen', [
      ['Code', 'Omschrijving'],
      ...acc.map(x => [x.code, x.naam]),
    ]);

    const begrotingRijen = [];
    Object.keys(bud).sort().forEach(jaar => (bud[jaar] || []).forEach(r => begrotingRijen.push([jaar, r.sectie, r.categorie, r.bedrag, r.notitie])));
    voegToe('Begroting', [['Boekjaar', 'Sectie', 'Categorie', 'Begroot bedrag', 'Notitie'], ...begrotingRijen]);

    voegToe('Boekjaren', [['Boekjaar'], ...bj.map(x => [x])]);

    voegToe('Vergaderingen', [
      ['Titel', 'Datum', 'Locatie', 'Aanwezigen', 'Aantal agendapunten'],
      ...verg.map(x => [x.titel, x.datum, x.locatie, x.aanwezigen, (x.agendapunten || []).length]),
    ]);

    const agendaRijen = [];
    verg.forEach(v => (v.agendapunten || []).forEach((p, idx) => agendaRijen.push([v.titel, v.datum, idx + 1, p.titel, p.toelichting, p.notulen])));
    voegToe('Agendapunten', [['Vergadering', 'Datum', 'Volgnr.', 'Titel', 'Toelichting', 'Notulen'], ...agendaRijen]);

    voegToe('Actielijst', [
      ['Omschrijving', 'Vergadering', 'Agendapunt', 'Wie', 'Deadline', 'Status'],
      ...act.map(x => [x.omschrijving, x.vergaderingTitel, x.agendapuntTitel, x.wie, x.deadline, x.status]),
    ]);

    voegToe('Contributies', [
      ['Lid', 'Boekjaar', 'Bedrag', 'Betaald', 'Datum'],
      ...contr.map(x => {
        const lid = m.find(y => y.id === x.lidId);
        return [lid ? fullName(lid) : x.lidId, x.jaar, x.bedrag, x.betaald ? 'ja' : 'nee', x.datum];
      }),
    ]);

    XLSX.writeFile(wb, `BladelsCreatief_Backup_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  const [herstelData, setHerstelData] = useState(null);
  const [herstelFout, setHerstelFout] = useState('');
  const [herstelResultaat, setHerstelResultaat] = useState(false);
  const [toonBevestiging, setToonBevestiging] = useState(false);

  function handleHerstelFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setHerstelFout(''); setHerstelResultaat(false); setHerstelData(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        if (!data || typeof data !== 'object' || !Array.isArray(data.members)) {
          setHerstelFout('Dit lijkt geen geldig BladelsCreatief-back-upbestand te zijn.');
          return;
        }
        setHerstelData(data);
      } catch (err) {
        setHerstelFout('Kon dit bestand niet lezen — is het een geldig .json back-upbestand?');
      }
    };
    reader.readAsText(file);
  }
  function herstelToepassen() {
    if (!herstelData) return;
    if (herstelData.members) setMembers(herstelData.members);
    if (herstelData.workshops) setWorkshops(herstelData.workshops);
    if (herstelData.inschrijvingen) setInschrijvingen(herstelData.inschrijvingen);
    if (herstelData.tx) setTx(herstelData.tx);
    if (herstelData.budget) setBudget(herstelData.budget);
    if (herstelData.boekjaren) setBoekjaren(herstelData.boekjaren);
    if (herstelData.vergaderingen) setVergaderingen(herstelData.vergaderingen);
    if (herstelData.actielijst) setActielijst(herstelData.actielijst);
    if (herstelData.contributies) setContributies(herstelData.contributies);
    if (herstelData.rolpermissies) setRolpermissies(herstelData.rolpermissies);
    if (herstelData.begrotingKoppelingen) setBegrotingKoppelingen(herstelData.begrotingKoppelingen);
    onLog(`Volledige back-up hersteld (bestand van ${herstelData.export_datum ? fmtDate(herstelData.export_datum) : 'onbekende datum'})`, 'instellingen');
    setToonBevestiging(false);
    setHerstelResultaat(true);
  }

  return (
    <div className="space-y-4">
      <div className="flex gap-1 flex-wrap">
        {SECTIES.map(s => {
          const Icon = s.icon;
          const active = sectie === s.id;
          return (
            <button key={s.id} onClick={() => setSectie(s.id)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-medium border"
              style={{ background: active ? C.clay : 'transparent', color: active ? '#fff' : C.ink, borderColor: active ? C.clay : C.border }}>
              <Icon size={14} /> {s.label}
            </button>
          );
        })}
      </div>

      {sectie === 'rollen' && (
        <RolBeheer rolpermissies={rolpermissies} setRolpermissies={setRolpermissies} readOnly={!isVoorzitter} onLog={onLog} />
      )}

      {sectie === 'beveiliging' && (
        <BeveiligingBeheer beveiliging={beveiliging} setBeveiliging={setBeveiliging} readOnly={!isVoorzitter} onLog={onLog} />
      )}

      {sectie === 'standaarden' && (
        <StandaardenBeheer standaarden={standaarden} setStandaarden={setStandaarden} readOnly={!isVoorzitter} onLog={onLog} />
      )}

      {sectie === 'grootboek' && (
        <GrootboekBeheer accounts={accounts} setAccounts={setAccounts} tx={tx} readOnly={!magFinancien} onLog={onLog} />
      )}

      {sectie === 'koppelingen' && (
        <BegrotingKoppelingenBeheer koppelingen={begrotingKoppelingen} setKoppelingen={setBegrotingKoppelingen} accounts={accounts} readOnly={!magFinancien} onLog={onLog} />
      )}

      {sectie === 'dagdelen' && (
        <DagdelenBeheer dagdelen={dagdelen} setDagdelen={setDagdelen} members={members} setMembers={setMembers} readOnly={!magLeden} onLog={onLog} />
      )}

      {sectie === 'soorten' && (
        <NaamlijstBeheer titel="workshop-soort" items={workshopSoorten} setItems={setWorkshopSoorten}
          isInGebruik={soort => workshops.some(w => w.soort === soort)}
          onRename={(oud, nieuw) => setWorkshops(workshops.map(w => w.soort === oud ? { ...w, soort: nieuw } : w))}
          readOnly={!magWorkshops}
          gebruiktLabel="workshops" onLog={onLog} logGebied="instellingen" />
      )}

      {sectie === 'agenda' && (
        <StandaardAgendaBeheer vooraf={agendapuntenVooraf} setVooraf={setAgendapuntenVooraf}
          afsluitend={agendapuntenAfsluitend} setAfsluitend={setAgendapuntenAfsluitend} readOnly={!magVergaderingen} onLog={onLog} />
      )}

      {sectie === 'logboek' && (
        <Card className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left border-b" style={{ borderColor: C.border }}>
                {['Tijdstip', 'Gebruiker', 'Gebied', 'Actie'].map(h => (
                  <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {logboek.map(e => (
                <tr key={e.id} className="border-b last:border-0" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2 whitespace-nowrap" style={{ color: C.inkSoft }}>{new Date(e.tijdstip).toLocaleString('nl-NL')}</td>
                  <td className="px-3 py-2 font-medium">{e.gebruiker}</td>
                  <td className="px-3 py-2"><Badge tone="muted">{e.gebied}</Badge></td>
                  <td className="px-3 py-2">{e.actie}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {!logboek.length && <EmptyState icon={History} text="Nog geen wijzigingen gelogd." />}
        </Card>
      )}

      {sectie === 'prullenbak' && (
        <Card className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left border-b" style={{ borderColor: C.border }}>
                {['Type', 'Naam/omschrijving', 'Verwijderd op', 'Door', ''].map(h => (
                  <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {prullenbak.map(e => (
                <tr key={e.id} className="border-b last:border-0" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2"><Badge tone="ochre">{e.type}</Badge></td>
                  <td className="px-3 py-2">{trashLabel(e)}</td>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{fmtDate(e.verwijderdOp)}</td>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{e.verwijderdDoor}</td>
                  <td className="px-3 py-2">
                    <div className="flex gap-2 justify-end">
                      <button onClick={() => herstel(e)} className="text-xs underline flex items-center gap-1" style={{ color: C.sageDeep }}><RotateCcw size={12} /> Herstellen</button>
                      <button onClick={() => verwijderDefinitief(e.id)} className="text-xs underline" style={{ color: C.rose }}>Definitief verwijderen</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {!prullenbak.length && <EmptyState icon={Trash2} text="Prullenbak is leeg." />}
        </Card>
      )}

      {sectie === 'importeren' && (
        <ImportSectie members={members} setMembers={setMembers} workshops={workshops} setWorkshops={setWorkshops}
          inschrijvingen={inschrijvingen} setInschrijvingen={setInschrijvingen}
          tx={tx} setTx={setTx} boekjaren={boekjaren} setBoekjaren={setBoekjaren} accounts={accounts} alleDagdelen={dagdelen}
          magLeden={magLedenImporteren} magWorkshops={magWorkshopsImporteren} magFinancien={magFinancienImporteren} onLog={onLog} />
      )}

      {sectie === 'backup' && (
        <div className="space-y-4">
          <Card className="p-5">
            <h3 className="font-semibold mb-2" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Volledige back-up</h3>
            <p className="text-sm mb-4" style={{ color: C.inkSoft }}>
              Download in één keer alle gegevens (leden, workshops, financiën, begroting, vergaderingen, actielijst en contributies) als reservekopie — je krijgt zowel een <strong>.json-bestand</strong> (voor het exact herstellen in deze app, zie hieronder) als een <strong>.xlsx-bestand</strong> met alle gegevens overzichtelijk in aparte Excel-tabbladen. Onmisbaar vlak vóórdat je een nieuwe versie van de app publiceert.
            </p>
            <div className="flex gap-2 flex-wrap">
              <Btn icon={DatabaseBackup} tone="sage" onClick={exportBackup}>Download volledige back-up (JSON + Excel)</Btn>
              <Btn icon={Download} tone="outline" onClick={exportBackupExcel}>Alleen Excel-bestand</Btn>
            </div>
            <p className="text-xs mt-2" style={{ color: C.inkSoft }}>Komt er bij de bovenste knop maar één bestand door? Gebruik dan de aparte Excel-knop hiernaast — sommige browsers blokkeren twee gelijktijdige downloads.</p>
          </Card>

          <Card className="p-5">
            <h3 className="font-semibold mb-2" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Back-up herstellen</h3>
            <p className="text-sm mb-1" style={{ color: C.inkSoft }}>
              Upload een eerder gedownload back-upbestand (.json) om alle gegevens hiermee te <strong>overschrijven</strong> — bijvoorbeeld nadat opnieuw publiceren van de app tot een lege/oude versie leidde.
            </p>
            {!isVoorzitter && (
              <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Alleen de Voorzitter kan een back-up herstellen — dit overschrijft namelijk alle gegevens in de hele app, over alle domeinen heen.</p>
            )}
            {isVoorzitter && (<>
            <p className="text-xs mb-3" style={{ color: C.rose }}>
              Let op: dit vervangt alle huidige gegevens in deze versie van de app door de inhoud van het back-upbestand. Deze actie kan niet ongedaan gemaakt worden.
            </p>
            {actieveAnderen == null && (
              <p className="text-xs mb-3" style={{ color: C.inkSoft }}>Controleren wie er nog meer actief is…</p>
            )}
            {actieveAnderen != null && actieveAnderen.length > 0 && (
              <div className="flex items-start gap-2 rounded-lg px-3 py-2 mb-3 text-sm" style={{ background: '#F3E0E0', color: C.rose }}>
                <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 2 }} />
                <span>Let op: {actieveAnderen.join(', ')} lijkt/lijken op dit moment ook actief te zijn in de app (activiteit in de laatste 3 minuten). Herstellen overschrijft mogelijk hun werk — overleg eerst, of wacht tot zij klaar zijn.</span>
              </div>
            )}
            {actieveAnderen != null && actieveAnderen.length === 0 && (
              <p className="text-xs mb-3 flex items-center gap-1.5" style={{ color: C.sageDeep }}><CheckCircle2 size={14} /> Er lijkt op dit moment niemand anders actief — veilig om te herstellen.</p>
            )}
            <button onClick={checkActieveGebruikers} className="text-xs underline mb-3 block" style={{ color: C.clay }}>ververs</button>
            <input type="file" accept=".json" onChange={handleHerstelFile} className="text-sm" />
            {herstelFout && <p className="text-xs mt-2" style={{ color: C.rose }}>{herstelFout}</p>}
            {herstelData && !herstelResultaat && (
              <div className="mt-3 rounded-lg px-3 py-2 text-sm" style={{ background: C.paperDim }}>
                <p className="mb-2">Gevonden in dit bestand: {herstelData.members?.length ?? 0} leden, {herstelData.workshops?.length ?? 0} workshops, {herstelData.inschrijvingen?.length ?? 0} inschrijvingen, {herstelData.tx?.length ?? 0} boekingen, {herstelData.vergaderingen?.length ?? 0} vergaderingen.</p>
                <Btn tone="dangerSolid" icon={RotateCcw} onClick={() => setToonBevestiging(true)}>Herstellen en huidige data overschrijven</Btn>
              </div>
            )}
            {herstelResultaat && (
              <div className="mt-3 rounded-lg px-3 py-2 text-sm flex items-center gap-2" style={{ background: C.paperDim, color: C.sageDeep }}>
                <CheckCircle2 size={15} /> Back-up hersteld.
              </div>
            )}
            </>)}
          </Card>
        </div>
      )}
      {toonBevestiging && (
        <ConfirmModal
          title="Back-up herstellen"
          message="Weet je zeker dat je alle huidige gegevens wilt overschrijven met dit back-upbestand? Dit kan niet ongedaan gemaakt worden."
          confirmLabel="Ja, overschrijven"
          onConfirm={herstelToepassen}
          onCancel={() => setToonBevestiging(false)}
        />
      )}
    </div>
  );
}

function parseExcelDatum(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v).trim();
}

function ImportSectie({ members, setMembers, workshops, setWorkshops, inschrijvingen, setInschrijvingen, tx, setTx, boekjaren, setBoekjaren, accounts, alleDagdelen, magLeden, magWorkshops, magFinancien, onLog }) {
  const eersteType = magLeden ? 'leden' : (magWorkshops ? 'jeugdatelier' : (magFinancien ? 'financien' : null));
  const [type, setType] = useState(eersteType);
  const [workbook, setWorkbook] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [sheetNaam, setSheetNaam] = useState('');
  const [rekening, setRekening] = useState('908');
  const [bestandsnaam, setBestandsnaam] = useState('');
  const [resultaat, setResultaat] = useState(null);
  const [fout, setFout] = useState('');

  if (!magLeden && !magWorkshops && !magFinancien) {
    return <EmptyState icon={Upload} text="Je hebt geen bewerkrechten voor Leden, Workshops of Financiën, en kunt daarom niets importeren." />;
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setBestandsnaam(file.name);
    setResultaat(null);
    setFout('');
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        setWorkbook(wb);
        setSheets(wb.SheetNames);
        let gok;
        if (type === 'leden') gok = wb.SheetNames.find(n => /leden/i.test(n));
        else if (type === 'jeugdatelier') gok = wb.SheetNames.find(n => /jeugd/i.test(n));
        else gok = wb.SheetNames.find(n => /908|319/.test(n));
        gok = gok || wb.SheetNames[0];
        setSheetNaam(gok);
        if (type === 'financien') {
          const m = /319/.test(gok) ? '319' : (/908/.test(gok) ? '908' : rekening);
          setRekening(m);
        }
      } catch (err) {
        setFout('Kon dit bestand niet lezen — is het een geldig Excel-bestand (.xlsx)?');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function rijNaarLid(row) {
    const dagdelen = alleDagdelen.filter(slot => row[slot] != null && row[slot] !== '');
    return {
      voornaam: row['Voornaam'] || '', tussenvoegsel: row['tussenvoegsel'] || '', achternaam: row['Achternaam'] || '',
      naam: row['Naam'] || row['Samengestelde naam'] || '',
      email: row['E-mailadres'] || '', adres: row['Adres'] || '', postcode: row['Postcode'] || '', woonplaats: row['Woonplaats'] || '',
      telefoon: row['Telefoonnr.'] || row['Telefoonnr'] || '',
      gebdatum: parseExcelDatum(row['Geb. datum']), lidsinds: parseExcelDatum(row['Lid sinds']), eindelidmaat: parseExcelDatum(row['Einde lidmaat']),
      status: row['Status lidmaatschap'] || 'actief', functie: row['Functie'] || '', dagdelen,
    };
  }
  function rijNaarInschrijving(row) {
    return {
      naam: row['Naam'] || '',
      email: row['E-mailadres ouder 1'] || '', email2: row['E-mailadres ouder 2'] || '',
      telefoon: row['Telefoonnr ouder 1'] || row['Telefoonnr. ouder 1'] || '', telefoon2: row['Telefoonnr ouder 2'] || row['Telefoonnr. ouder 2'] || '',
      leeftijd: row['Leeftijd'] || null, groep: row['Groep'] || null,
      blok: row['Blok'] || 'Onbekend blok', betaald: !!row['Betaald'], datumInschrijving: parseExcelDatum(row['inschrijving'] || row['Inschrijving']),
    };
  }

  function importeerLeden(rows) {
    let nieuw = 0, bijgewerkt = 0;
    const volgende = [...members];
    rows.forEach(row => {
      const data = rijNaarLid(row);
      const key = fullName(data).trim().toLowerCase();
      if (!key) return;
      const idx = volgende.findIndex(m => fullName(m).trim().toLowerCase() === key);
      if (idx >= 0) { volgende[idx] = { ...volgende[idx], ...data }; bijgewerkt++; }
      else { volgende.push({ ...data, id: uid(volgende) }); nieuw++; }
    });
    setMembers(volgende);
    onLog(`Leden geïmporteerd uit Excel (${bestandsnaam}): ${nieuw} nieuw, ${bijgewerkt} bijgewerkt`, 'instellingen');
    return { nieuw, bijgewerkt };
  }
  function importeerJeugdatelier(rows) {
    let nieuweWorkshops = 0, nieuweInschr = 0, bijgewerkt = 0;
    const volgendeWorkshops = [...workshops];
    const volgendeInschr = [...inschrijvingen];
    rows.forEach(row => {
      const { blok, ...data } = rijNaarInschrijving(row);
      if (!data.naam) return;
      let ws = volgendeWorkshops.find(w => w.soort === 'Jeugdatelier' && w.titel.trim().toLowerCase() === blok.trim().toLowerCase());
      if (!ws) {
        ws = { id: uid(volgendeWorkshops), titel: blok, soort: 'Jeugdatelier', type: 'reeks', datums: [], dagdeel: '', locatie: '', bedrag: null, maxDeelnemers: null, status: 'open' };
        volgendeWorkshops.push(ws);
        nieuweWorkshops++;
      }
      const idx = volgendeInschr.findIndex(i => i.workshopId === ws.id && (i.naam || '').trim().toLowerCase() === data.naam.trim().toLowerCase());
      if (idx >= 0) { volgendeInschr[idx] = { ...volgendeInschr[idx], ...data }; bijgewerkt++; }
      else { volgendeInschr.push({ id: uid(volgendeInschr), workshopId: ws.id, herkomst: 'extern', lidId: null, status: 'ingeschreven', bedrag: null, notitie: '', ...data }); nieuweInschr++; }
    });
    setWorkshops(volgendeWorkshops);
    setInschrijvingen(volgendeInschr);
    onLog(`Jeugdatelier geïmporteerd uit Excel (${bestandsnaam}): ${nieuweWorkshops} nieuwe workshops, ${nieuweInschr} nieuwe inschrijvingen, ${bijgewerkt} bijgewerkt`, 'instellingen');
    return { nieuweWorkshops, nieuweInschr, bijgewerkt };
  }

  function rijenNaarTransacties(ws) {
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    let start = null;
    for (let i = 0; i < raw.length; i++) {
      if (raw[i][0] === 'Maand') { start = i + 1; break; }
    }
    if (start == null) return [];
    const result = [];
    for (let i = start; i < raw.length; i++) {
      const [maand, datum, grootboek, bedrag] = raw[i];
      const omschrijving = raw[i][4];
      if (!datum || grootboek == null || bedrag == null) continue;
      const gbStr = String(grootboek).trim();
      const gbCode = gbStr.split(/\s+/)[0];
      const gbNaam = gbStr.slice(gbCode.length).trim();
      const datumStr = parseExcelDatum(datum);
      const datumObj = new Date(datumStr);
      result.push({
        rekening, jaar: isNaN(datumObj) ? new Date().getFullYear() : datumObj.getFullYear(),
        maand: Number(maand) || (isNaN(datumObj) ? 1 : datumObj.getMonth() + 1),
        datum: datumStr, grootboek_code: gbCode, grootboek_naam: gbNaam, bedrag: Number(bedrag), omschrijving: omschrijving || '',
      });
    }
    return result;
  }
  function importeerFinancien(rows) {
    let nieuw = 0, bijgewerkt = 0;
    const volgende = [...tx];
    const jaren = new Set(boekjaren);
    rows.forEach(data => {
      const idx = volgende.findIndex(t => t.rekening === data.rekening && t.datum === data.datum && t.grootboek_code === data.grootboek_code && (t.omschrijving || '') === (data.omschrijving || ''));
      if (idx >= 0) { volgende[idx] = { ...volgende[idx], ...data }; bijgewerkt++; }
      else { volgende.push({ ...data, id: uid(volgende) }); nieuw++; }
      jaren.add(data.jaar);
    });
    setTx(volgende);
    if (jaren.size !== boekjaren.length) setBoekjaren(Array.from(jaren).sort((a, b) => a - b));
    onLog(`Financiën geïmporteerd uit Excel (${bestandsnaam}, rekening .${rekening}): ${nieuw} nieuw, ${bijgewerkt} bijgewerkt`, 'instellingen');
    return { nieuw, bijgewerkt };
  }
  function verwerk() {
    if (!workbook || !sheetNaam) return;
    const ws = workbook.Sheets[sheetNaam];
    if (type === 'financien') {
      setResultaat(importeerFinancien(rijenNaarTransacties(ws)));
      return;
    }
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
    setResultaat(type === 'leden' ? importeerLeden(rows) : importeerJeugdatelier(rows));
  }

  return (
    <Card className="p-5 space-y-4">
      <div>
        <span className="block text-xs font-medium mb-1.5" style={{ color: C.inkSoft }}>Wat wil je importeren?</span>
        <div className="flex gap-1 flex-wrap">
          {magLeden && (
            <button onClick={() => { setType('leden'); setResultaat(null); }} className="px-3 py-1.5 rounded-lg text-sm font-medium border"
              style={{ background: type === 'leden' ? C.clay : 'transparent', color: type === 'leden' ? '#fff' : C.ink, borderColor: type === 'leden' ? C.clay : C.border }}>
              Leden
            </button>
          )}
          {magWorkshops && (
            <button onClick={() => { setType('jeugdatelier'); setResultaat(null); }} className="px-3 py-1.5 rounded-lg text-sm font-medium border"
              style={{ background: type === 'jeugdatelier' ? C.clay : 'transparent', color: type === 'jeugdatelier' ? '#fff' : C.ink, borderColor: type === 'jeugdatelier' ? C.clay : C.border }}>
              Jeugdatelier / workshop-inschrijvingen
            </button>
          )}
          {magFinancien && (
            <button onClick={() => { setType('financien'); setResultaat(null); }} className="px-3 py-1.5 rounded-lg text-sm font-medium border"
              style={{ background: type === 'financien' ? C.clay : 'transparent', color: type === 'financien' ? '#fff' : C.ink, borderColor: type === 'financien' ? C.clay : C.border }}>
              Financiën (boekingen)
            </button>
          )}
        </div>
      </div>
      <div>
        <span className="block text-xs font-medium mb-1.5" style={{ color: C.inkSoft }}>Excel-bestand (.xlsx)</span>
        <input type="file" accept=".xlsx,.xls" onChange={handleFile} className="text-sm" />
        {bestandsnaam && <p className="text-xs mt-1" style={{ color: C.inkSoft }}>Geselecteerd: {bestandsnaam}</p>}
        {fout && <p className="text-xs mt-1" style={{ color: C.rose }}>{fout}</p>}
      </div>
      {sheets.length > 0 && (
        <Field label="Werkblad (tabblad) in dit bestand">
          <select className={inputCls} style={inputStyle} value={sheetNaam} onChange={e => setSheetNaam(e.target.value)}>
            {sheets.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
      )}
      {type === 'financien' && (
        <Field label="Op welke rekening staan deze boekingen?">
          <select className={inputCls} style={inputStyle} value={rekening} onChange={e => setRekening(e.target.value)}>
            <option value="908">.908 lopend</option>
            <option value="319">.319 spaar</option>
          </select>
        </Field>
      )}
      <p className="text-xs" style={{ color: C.inkSoft }}>
        {type === 'leden' && 'Bestaande leden (op volledige naam) worden bijgewerkt met de gegevens uit dit bestand; nieuwe namen worden toegevoegd. Er wordt niets verwijderd.'}
        {type === 'jeugdatelier' && 'Bestaande inschrijvingen (op naam + blok) worden bijgewerkt; nieuwe namen of nieuwe blokken (als workshop) worden toegevoegd. Er wordt niets verwijderd.'}
        {type === 'financien' && 'Verwacht hetzelfde format als de bestaande financiële administratie (kolommen Maand/Datum/Grootboek/Bedrag/Omschrijving, met "Maand" als koprij). Boekingen die al bestaan (zelfde datum, grootboek en omschrijving) worden bijgewerkt; nieuwe boekingen worden toegevoegd. Er wordt niets verwijderd.'}
      </p>
      <Btn icon={Upload} onClick={verwerk} disabled={!workbook || !sheetNaam}>Importeren</Btn>
      {resultaat && (
        <div className="rounded-lg px-3 py-2 text-sm flex items-center gap-2" style={{ background: C.paperDim, color: C.sageDeep }}>
          <CheckCircle2 size={15} />
          {type === 'leden' && <span>{resultaat.nieuw} nieuwe leden toegevoegd, {resultaat.bijgewerkt} bestaande leden bijgewerkt.</span>}
          {type === 'jeugdatelier' && <span>{resultaat.nieuweWorkshops} nieuwe workshops, {resultaat.nieuweInschr} nieuwe inschrijvingen toegevoegd, {resultaat.bijgewerkt} bestaande inschrijvingen bijgewerkt.</span>}
          {type === 'financien' && <span>{resultaat.nieuw} nieuwe boekingen toegevoegd, {resultaat.bijgewerkt} bestaande boekingen bijgewerkt.</span>}
        </div>
      )}
    </Card>
  );
}

function NaamlijstBeheer({ titel, items, setItems, isInGebruik, gebruiktLabel, onRename, readOnly, onLog, logGebied }) {
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [delItem, setDelItem] = useState(null);
  const [fout, setFout] = useState('');

  function opslaan(nieuweNaam) {
    const n = nieuweNaam.trim();
    if (!n) { setFout(`Vul een naam in.`); return; }
    if (editing == null) {
      if (items.includes(n)) { setFout(`"${n}" bestaat al.`); return; }
      setItems([...items, n]);
      onLog(`${titel} toegevoegd: ${n}`, logGebied);
    } else {
      setItems(items.map(x => x === editing ? n : x));
      if (onRename && editing !== n) onRename(editing, n);
      onLog(`${titel} bewerkt: ${editing} → ${n}`, logGebied);
    }
    setFout('');
    setShowForm(false);
    setEditing(null);
  }
  function verwijderen(item) {
    setItems(items.filter(x => x !== item));
    onLog(`${titel} verwijderd: ${item}`, logGebied);
  }

  return (
    <div className="space-y-3">
      {readOnly && <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Je hebt geen bewerkrechten voor Workshops, en kunt workshop-soorten daarom alleen bekijken.</p>}
      <div className="flex items-center justify-between">
        <p className="text-xs" style={{ color: C.inkSoft }}>{items.length} {titel === 'workshop-soort' ? "workshop-soorten" : titel}</p>
        {!readOnly && <Btn icon={Plus} onClick={() => { setEditing(null); setShowForm(true); setFout(''); }}>{titel === 'workshop-soort' ? 'Workshop-soort toevoegen' : 'Toevoegen'}</Btn>}
      </div>
      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Naam</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {items.map(item => {
              const geblokkeerd = isInGebruik(item);
              return (
                <tr key={item} className="border-b last:border-0 hover:bg-black/[0.02]" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2">{item}</td>
                  <td className="px-3 py-2">
                    {!readOnly && (
                      <div className="flex gap-1 justify-end">
                        <button onClick={() => { setEditing(item); setShowForm(true); setFout(''); }}
                          title="Bewerken (ook mogelijk als dit al bij iets is gekoppeld)"
                          className="p-1.5 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={14} /></button>
                        <button onClick={() => geblokkeerd ? null : setDelItem(item)}
                          title={geblokkeerd ? `In gebruik bij ${gebruiktLabel} — kan niet verwijderd worden` : 'Verwijderen'}
                          className="p-1.5 rounded hover:bg-black/5" style={{ color: geblokkeerd ? '#c9c2b3' : C.rose, cursor: geblokkeerd ? 'not-allowed' : 'pointer' }}><Trash2 size={14} /></button>
                      </div>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {!items.length && <EmptyState icon={Palette} text="Nog niets toegevoegd." />}
      </Card>
      {showForm && (
        <PromptModal title={editing ? `${titel} bewerken` : `${titel} toevoegen`} label="Naam" placeholder="Naam"
          onSave={opslaan} onCancel={() => { setShowForm(false); setEditing(null); }} />
      )}
      {fout && <p className="text-xs" style={{ color: C.rose }}>{fout}</p>}
      {delItem != null && (
        <ConfirmModal message={`"${delItem}" verwijderen?`} onConfirm={() => { verwijderen(delItem); setDelItem(null); }} onCancel={() => setDelItem(null)} />
      )}
    </div>
  );
}

function StandaardAgendaBeheer({ vooraf, setVooraf, afsluitend, setAfsluitend, readOnly, onLog }) {
  function Lijst({ label, items, setItems }) {
    const [showForm, setShowForm] = useState(false);

    function voegToe(naam) {
      const n = naam.trim();
      if (!n) return;
      setItems([...items, n]);
      onLog(`Standaard agendapunt toegevoegd bij "${label}": ${n}`, 'instellingen');
      setShowForm(false);
    }
    function verwijderen(idx) {
      onLog(`Standaard agendapunt verwijderd bij "${label}": ${items[idx]}`, 'instellingen');
      setItems(items.filter((_, i) => i !== idx));
    }
    function move(idx, richting) {
      const nieuw = [...items];
      const doel = idx + richting;
      if (doel < 0 || doel >= nieuw.length) return;
      [nieuw[idx], nieuw[doel]] = [nieuw[doel], nieuw[idx]];
      setItems(nieuw);
    }

    return (
      <Card className="p-4">
        <div className="flex items-center justify-between mb-2">
          <h3 className="font-semibold text-sm" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>{label}</h3>
          {!readOnly && <Btn icon={Plus} onClick={() => setShowForm(true)}>Agendapunt toevoegen</Btn>}
        </div>
        <ul className="space-y-1">
          {items.map((item, i) => (
            <li key={i} className="flex items-center gap-2 px-2.5 py-1.5 rounded-lg" style={{ background: C.paperDim }}>
              <span className="flex-1 text-sm">{item}</span>
              {!readOnly && (
                <>
                  <button onClick={() => move(i, -1)} disabled={i === 0} className="p-0.5 rounded hover:bg-black/5 disabled:opacity-20" style={{ color: C.inkSoft }}><ChevronUp size={14} /></button>
                  <button onClick={() => move(i, 1)} disabled={i === items.length - 1} className="p-0.5 rounded hover:bg-black/5 disabled:opacity-20" style={{ color: C.inkSoft }}><ChevronDown size={14} /></button>
                  <button onClick={() => verwijderen(i)} className="p-1 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={14} /></button>
                </>
              )}
            </li>
          ))}
          {!items.length && <li className="text-xs italic" style={{ color: C.inkSoft }}>Nog niets toegevoegd.</li>}
        </ul>
        {showForm && (
          <PromptModal title={`Agendapunt toevoegen bij "${label}"`} label="Titel" placeholder="bv. Opening"
            onSave={voegToe} onCancel={() => setShowForm(false)} />
        )}
      </Card>
    );
  }

  return (
    <div className="space-y-4">
      <p className="text-xs" style={{ color: C.inkSoft }}>
        Deze punten worden als aanvinkoptie voorgesteld bij het aanmaken van een nieuwe vergadering ("Standaard agendapunten toevoegen"). "Vooraf" komt bovenaan de agenda, "Afsluitend" onderaan.
      </p>
      {readOnly && <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Je hebt geen bewerkrechten voor Vergaderingen, en kunt dit daarom alleen bekijken.</p>}
      <Lijst label="Vooraf" items={vooraf} setItems={setVooraf} />
      <Lijst label="Afsluitend" items={afsluitend} setItems={setAfsluitend} />
    </div>
  );
}

function DagdelenBeheer({ dagdelen, setDagdelen, members, setMembers, readOnly, onLog }) {
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [delSlot, setDelSlot] = useState(null);
  const [fout, setFout] = useState('');

  function ledenInGroep(slot) {
    return members.filter(m => (m.dagdelen || []).includes(slot));
  }

  function opslaan(nieuweNaam) {
    const n = nieuweNaam.trim();
    if (!n) { setFout('Vul een naam voor het dagdeel/groep in.'); return; }
    if (editing == null) {
      if (dagdelen.includes(n)) { setFout('Dit dagdeel/deze groep bestaat al.'); return; }
      setDagdelen([...dagdelen, n]);
      onLog(`Dagdeel/groep toegevoegd: ${n}`, 'instellingen');
    } else {
      setDagdelen(dagdelen.map(d => d === editing ? n : d));
      if (editing !== n) {
        setMembers(members.map(m => (m.dagdelen || []).includes(editing)
          ? { ...m, dagdelen: m.dagdelen.map(d => d === editing ? n : d) }
          : m));
      }
      onLog(`Dagdeel/groep bewerkt: ${editing} → ${n}`, 'instellingen');
    }
    setFout('');
    setShowForm(false);
    setEditing(null);
  }
  function verwijderen(slot) {
    setDagdelen(dagdelen.filter(d => d !== slot));
    onLog(`Dagdeel/groep verwijderd: ${slot}`, 'instellingen');
  }

  return (
    <div className="space-y-3">
      {readOnly && <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Je hebt geen bewerkrechten voor Leden, en kunt dagdelen/groepen daarom alleen bekijken.</p>}
      <div className="flex items-center justify-between">
        <p className="text-xs" style={{ color: C.inkSoft }}>{dagdelen.length} dagdelen/groepen</p>
        {!readOnly && <Btn icon={Plus} onClick={() => { setEditing(null); setShowForm(true); setFout(''); }}>Dagdeel/groep toevoegen</Btn>}
      </div>
      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Dagdeel / groep</th>
              <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Leden</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {dagdelen.map(slot => {
              const leden = ledenInGroep(slot);
              const geblokkeerd = leden.length > 0;
              return (
                <tr key={slot} className="border-b last:border-0 hover:bg-black/[0.02]" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2">{slot}</td>
                  <td className="px-3 py-2"><Badge tone={geblokkeerd ? 'clay' : 'muted'}>{leden.length}</Badge></td>
                  <td className="px-3 py-2">
                    {!readOnly && (
                      <div className="flex gap-1 justify-end">
                        <button onClick={() => { setEditing(slot); setShowForm(true); setFout(''); }}
                          title="Bewerken (ook mogelijk als hier al leden aan gekoppeld zijn)"
                          className="p-1.5 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={14} /></button>
                        <button onClick={() => geblokkeerd ? null : setDelSlot(slot)}
                          title={geblokkeerd ? 'Er zijn leden aan deze groep toegevoegd — kan niet verwijderd worden' : 'Verwijderen'}
                          className="p-1.5 rounded hover:bg-black/5" style={{ color: geblokkeerd ? '#c9c2b3' : C.rose, cursor: geblokkeerd ? 'not-allowed' : 'pointer' }}><Trash2 size={14} /></button>
                      </div>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {!dagdelen.length && <EmptyState icon={Grid3x3} text="Nog geen dagdelen/groepen." />}
      </Card>
      {showForm && (
        <PromptModal title={editing ? 'Dagdeel/groep bewerken' : 'Dagdeel/groep toevoegen'} label="Naam" placeholder="bv. vr 10.00 - 13.00"
          onSave={opslaan} onCancel={() => { setShowForm(false); setEditing(null); }} />
      )}
      {fout && <p className="text-xs" style={{ color: C.rose }}>{fout}</p>}
      {delSlot != null && (
        <ConfirmModal message={`Dagdeel/groep "${delSlot}" verwijderen?`} onConfirm={() => { verwijderen(delSlot); setDelSlot(null); }} onCancel={() => setDelSlot(null)} />
      )}
    </div>
  );
}

function BegrotingKoppelingenBeheer({ koppelingen, setKoppelingen, accounts, readOnly, onLog }) {
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [delId, setDelId] = useState(null);

  function opslaan(data) {
    if (data.id) {
      setKoppelingen(koppelingen.map(k => k.id === data.id ? data : k));
      onLog(`Begroting-koppeling bewerkt: ${data.trefwoord}`, 'instellingen');
    } else {
      setKoppelingen([...koppelingen, { ...data, id: uid(koppelingen) }]);
      onLog(`Begroting-koppeling toegevoegd: ${data.trefwoord}`, 'instellingen');
    }
    setShowForm(false);
    setEditing(null);
  }
  function verwijderen(id) {
    const k = koppelingen.find(x => x.id === id);
    setKoppelingen(koppelingen.filter(x => x.id !== id));
    if (k) onLog(`Begroting-koppeling verwijderd: ${k.trefwoord}`, 'instellingen');
  }

  return (
    <div className="space-y-3">
      <p className="text-xs" style={{ color: C.inkSoft }}>
        Bepaalt hoe "werkelijk"-bedragen in Begroting/Rapportage worden gevonden voor categorieën die niet met een grootboekcode beginnen (bijvoorbeeld "Contributie"). Begint een categorie al met een code (bv. "4300 Huur"), dan wordt die rechtstreeks gebruikt en is deze tabel niet nodig. Rijen worden van boven naar beneden doorzocht — de eerste match op trefwoord (ergens in de categorienaam, niet hoofdlettergevoelig) telt.
      </p>
      {readOnly && <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Je hebt geen bewerkrechten voor Financiën, en kunt deze koppelingen daarom alleen bekijken.</p>}
      <div className="flex items-center justify-between">
        <p className="text-xs" style={{ color: C.inkSoft }}>{koppelingen.length} koppelingen</p>
        {!readOnly && <Btn icon={Plus} onClick={() => { setEditing(null); setShowForm(true); }}>Koppeling toevoegen</Btn>}
      </div>
      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Trefwoord</th>
              <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Grootboekrekening(en)</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {koppelingen.map(k => (
              <tr key={k.id} className="border-b last:border-0 hover:bg-black/[0.02]" style={{ borderColor: C.border }}>
                <td className="px-3 py-2">{k.trefwoord}</td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>{k.uitsluiten ? <Badge tone="rose">niet meetellen</Badge> : (k.codes || '—')}</td>
                <td className="px-3 py-2">
                  {!readOnly && (
                    <div className="flex gap-1 justify-end">
                      <button onClick={() => { setEditing(k); setShowForm(true); }} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={14} /></button>
                      <button onClick={() => setDelId(k.id)} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={14} /></button>
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!koppelingen.length && <EmptyState icon={Link2} text="Nog geen koppelingen." />}
      </Card>
      {showForm && (
        <BegrotingKoppelingModal item={editing} accounts={accounts} onSave={opslaan} onClose={() => { setShowForm(false); setEditing(null); }} />
      )}
      {delId != null && (
        <ConfirmModal message="Deze koppeling verwijderen?" onConfirm={() => { verwijderen(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
    </div>
  );
}

function BegrotingKoppelingModal({ item, accounts, onSave, onClose }) {
  const [trefwoord, setTrefwoord] = useState(item ? item.trefwoord : '');
  const [codes, setCodes] = useState(item ? item.codes : '');
  const [uitsluiten, setUitsluiten] = useState(item ? item.uitsluiten : false);
  return (
    <Modal title={item ? 'Koppeling bewerken' : 'Koppeling toevoegen'} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Trefwoord (herkend als onderdeel van de categorienaam)"><input className={inputCls} style={inputStyle} placeholder="bv. contributie" value={trefwoord} onChange={e => setTrefwoord(e.target.value)} /></Field>
        <label className="flex items-center gap-2 text-sm">
          <input type="checkbox" checked={uitsluiten} onChange={e => setUitsluiten(e.target.checked)} />
          Deze categorie nooit laten meetellen bij "werkelijk" (voor rekenregels zoals "Resultaat")
        </label>
        {!uitsluiten && (
          <Field label="Grootboekrekening(en) — code(s), gescheiden door een komma bij meerdere">
            <input className={inputCls} style={inputStyle} placeholder="bv. 8002,8005" value={codes} onChange={e => setCodes(e.target.value)} />
            <p className="text-xs mt-1" style={{ color: C.inkSoft }}>
              {codes.split(',').map(c => c.trim()).filter(Boolean).map(c => {
                const acc = accounts.find(a => a.code === c);
                return acc ? `${c} (${acc.naam})` : `${c} (onbekende code)`;
              }).join(', ') || 'Nog geen code ingevuld.'}
            </p>
          </Field>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave({ id: item ? item.id : null, trefwoord: trefwoord.trim(), codes: uitsluiten ? '' : codes.trim(), uitsluiten })} disabled={!trefwoord.trim()}>{item ? 'Opslaan' : 'Toevoegen'}</Btn>
      </div>
    </Modal>
  );
}

function GrootboekBeheer({ accounts, setAccounts, tx, readOnly, onLog }) {
  const [editing, setEditing] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [delCode, setDelCode] = useState(null);
  const [fout, setFout] = useState('');
  const gesorteerd = [...accounts].sort((a, b) => a.code.localeCompare(b.code));

  function heeftBoekingen(code) {
    return tx.some(t => t.grootboek_code === code);
  }

  function opslaan(data) {
    const c = data.code.trim();
    const n = data.naam.trim();
    if (!c || !n) { setFout('Vul zowel een code als een omschrijving in.'); return; }
    if (!editing && accounts.some(a => a.code === c)) { setFout(`Grootboekrekening ${c} bestaat al.`); return; }
    if (editing) {
      setAccounts(accounts.map(a => a.code === editing.code ? { code: c, naam: n } : a));
      onLog(`Grootboekrekening bewerkt: ${c} ${n}`, 'instellingen');
    } else {
      setAccounts([...accounts, { code: c, naam: n }]);
      onLog(`Grootboekrekening toegevoegd: ${c} ${n}`, 'instellingen');
    }
    setFout('');
    setShowForm(false);
    setEditing(null);
  }
  function verwijderen(code) {
    const acc = accounts.find(a => a.code === code);
    setAccounts(accounts.filter(a => a.code !== code));
    if (acc) onLog(`Grootboekrekening verwijderd: ${acc.code} ${acc.naam}`, 'instellingen');
  }

  return (
    <div className="space-y-3">
      {readOnly && <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Je hebt geen bewerkrechten voor Financiën, en kunt grootboekrekeningen daarom alleen bekijken.</p>}
      <div className="flex items-center justify-between">
        <p className="text-xs" style={{ color: C.inkSoft }}>{accounts.length} grootboekrekeningen</p>
        {!readOnly && <Btn icon={Plus} onClick={() => { setEditing(null); setShowForm(true); setFout(''); }}>Grootboekrekening toevoegen</Btn>}
      </div>
      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Code</th>
              <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Omschrijving</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {gesorteerd.map(a => {
              const geboekt = heeftBoekingen(a.code);
              return (
                <tr key={a.code} className="border-b last:border-0 hover:bg-black/[0.02]" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2" style={{ color: C.inkSoft }}>{a.code}</td>
                  <td className="px-3 py-2">{a.naam}</td>
                  <td className="px-3 py-2">
                    {!readOnly && (
                      <div className="flex gap-1 justify-end">
                        <button onClick={() => { setEditing(a); setShowForm(true); setFout(''); }} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={14} /></button>
                        <button onClick={() => geboekt ? null : setDelCode(a.code)} title={geboekt ? 'Er staan boekingen op deze rekening — kan niet verwijderd worden' : 'Verwijderen'}
                          className="p-1.5 rounded hover:bg-black/5" style={{ color: geboekt ? '#c9c2b3' : C.rose, cursor: geboekt ? 'not-allowed' : 'pointer' }}><Trash2 size={14} /></button>
                      </div>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {!accounts.length && <EmptyState icon={Wallet} text="Nog geen grootboekrekeningen." />}
      </Card>
      {showForm && (
        <GrootboekNieuwModal item={editing} onSave={opslaan} onClose={() => { setShowForm(false); setEditing(null); }} fout={fout} />
      )}
      {delCode != null && (
        <ConfirmModal message={`Grootboekrekening ${delCode} verwijderen?`} onConfirm={() => { verwijderen(delCode); setDelCode(null); }} onCancel={() => setDelCode(null)} />
      )}
    </div>
  );
}

function GrootboekNieuwModal({ item, onSave, onClose, fout }) {
  const [code, setCode] = useState(item ? item.code : '');
  const [naam, setNaam] = useState(item ? item.naam : '');
  return (
    <Modal title={item ? 'Grootboekrekening bewerken' : 'Grootboekrekening toevoegen'} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Code"><input disabled={!!item} className={inputCls} style={inputStyle} placeholder="bv. 8025" value={code} onChange={e => setCode(e.target.value)} /></Field>
        <Field label="Omschrijving"><input className={inputCls} style={inputStyle} placeholder="bv. Sponsorbijdragen" value={naam} onChange={e => setNaam(e.target.value)} /></Field>
        {fout && <p className="text-xs" style={{ color: C.rose }}>{fout}</p>}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave({ code, naam })} disabled={!code.trim() || !naam.trim()}>{item ? 'Opslaan' : 'Toevoegen'}</Btn>
      </div>
    </Modal>
  );
}

function StandaardenBeheer({ standaarden, setStandaarden, readOnly, onLog }) {
  function bijwerken(patch) {
    setStandaarden({ ...standaarden, ...patch });
    onLog(`Standaarden gewijzigd: ${Object.entries(patch).map(([k, v]) => `${k}=${v}`).join(', ')}`, 'instellingen');
  }
  return (
    <div className="space-y-4">
      {readOnly && <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Alleen de Voorzitter kan deze standaarden aanpassen.</p>}
      <Card className="p-5">
        <h3 className="font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Automatisch uitloggen</h3>
        <p className="text-sm mb-3" style={{ color: C.inkSoft }}>
          Hoelang iemand zonder muis-/toetsenbord-/touch-activiteit ingelogd mag blijven voordat automatisch wordt uitgelogd. Zet dit ruimer als de app bijvoorbeeld ook wordt gebruikt om tijdens een vergadering mee te lezen zonder steeds te typen.
        </p>
        <Field label="Uitloggen na inactiviteit van">
          <select disabled={readOnly} className={inputCls} style={inputStyle} value={standaarden.idleTimeoutMinuten} onChange={e => bijwerken({ idleTimeoutMinuten: Number(e.target.value) })}>
            <option value={3}>3 minuten</option>
            <option value={5}>5 minuten</option>
            <option value={10}>10 minuten</option>
            <option value={15}>15 minuten</option>
            <option value={30}>30 minuten</option>
            <option value={60}>60 minuten</option>
          </select>
        </Field>
      </Card>

      <Card className="p-5">
        <h3 className="font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Logogrootte in Word-exports</h3>
        <p className="text-sm mb-3" style={{ color: C.inkSoft }}>
          De hoogte van het BladelsCreatief-logo bovenaan geëxporteerde Word-documenten (notulen, actielijst, financieel verslag, jaarverslag).
        </p>
        <Field label="Logohoogte">
          <select disabled={readOnly} className={inputCls} style={inputStyle} value={standaarden.logoHoogteCm} onChange={e => bijwerken({ logoHoogteCm: Number(e.target.value) })}>
            <option value={1.5}>1,5 cm</option>
            <option value={2}>2 cm</option>
            <option value={2.5}>2,5 cm</option>
            <option value={3}>3 cm</option>
            <option value={4}>4 cm</option>
            <option value={5}>5 cm</option>
          </select>
        </Field>
      </Card>
    </div>
  );
}
function BeveiligingBeheer({ beveiliging, setBeveiliging, readOnly, onLog }) {
  function bijwerken(patch) {
    setBeveiliging({ ...beveiliging, ...patch });
    onLog(`Beveiligingsinstellingen gewijzigd: ${Object.entries(patch).map(([k, v]) => `${k}=${v}`).join(', ')}`, 'instellingen');
  }
  return (
    <div className="space-y-4">
      <Card className="p-5">
        <h3 className="font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Twee-factor authenticatie (2FA)</h3>
        <p className="text-sm mb-3" style={{ color: C.inkSoft }}>
          Vraagt bij het inloggen, naast de pincode, om een 6-cijferige code uit een authenticator-app (bv. Google Authenticator). Staat standaard uit.
        </p>
        <label className="flex items-center gap-2 text-sm mb-3">
          <input type="checkbox" disabled={readOnly} checked={beveiliging.tfaVerplicht} onChange={e => bijwerken({ tfaVerplicht: e.target.checked })} />
          Twee-factor authenticatie verplicht stellen
        </label>
        <Field label="Vertrouwensperiode — hoelang iemand na een geslaagde 2FA-controle niet opnieuw om de code wordt gevraagd">
          <select disabled={readOnly || !beveiliging.tfaVerplicht} className={inputCls} style={inputStyle}
            value={beveiliging.tfaVertrouwensdagen} onChange={e => bijwerken({ tfaVertrouwensdagen: Number(e.target.value) })}>
            <option value={7}>7 dagen</option>
            <option value={30}>30 dagen</option>
          </select>
        </Field>
        <p className="text-xs mt-2" style={{ color: C.inkSoft }}>
          Let op: deze vertrouwensperiode geldt per persoon (via de gedeelde opslag van de app), niet per specifiek apparaat — echte apparaatherkenning is in deze omgeving niet mogelijk.
        </p>
      </Card>

      <Card className="p-5">
        <h3 className="font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Wachtwoordeisen</h3>
        <p className="text-sm mb-3" style={{ color: C.inkSoft }}>
          Schakel dit pas in als de testfase is afgerond. Bij inschakelen moet een nieuw of gewijzigd wachtwoord minimaal 8 tekens bevatten, met een hoofdletter, kleine letter, cijfer en bijzonder teken. Al ingestelde eenvoudige pincodes blijven werken totdat iemand die zelf opnieuw instelt.
        </p>
        <label className="flex items-center gap-2 text-sm">
          <input type="checkbox" disabled={readOnly} checked={beveiliging.wachtwoordEisen} onChange={e => bijwerken({ wachtwoordEisen: e.target.checked })} />
          Wachtwoordeisen inschakelen (in plaats van eenvoudige pincode)
        </label>
      </Card>
    </div>
  );
}

function RolBeheer({ rolpermissies, setRolpermissies, readOnly, onLog }) {
  const [showNew, setShowNew] = useState(false);

  function updateRegel(id, patch) {
    setRolpermissies(rolpermissies.map(r => r.id === id ? { ...r, ...patch } : r));
  }
  function toggleTab(regel, tabId) {
    const has = regel.tabs.includes(tabId);
    updateRegel(regel.id, { tabs: has ? regel.tabs.filter(t => t !== tabId) : [...regel.tabs, tabId] });
  }
  function removeRegel(id) {
    const r = rolpermissies.find(x => x.id === id);
    setRolpermissies(rolpermissies.filter(x => x.id !== id));
    if (r) onLog(`Rol verwijderd: ${r.patroon}`, 'instellingen');
  }
  function addRegel(patroon) {
    setRolpermissies([...rolpermissies, { id: uid(rolpermissies), patroon, tabs: [] }]);
    onLog(`Rol toegevoegd: ${patroon}`, 'instellingen');
    setShowNew(false);
  }

  return (
    <div className="space-y-3">
      <p className="text-xs" style={{ color: C.inkSoft }}>
        Een bestuurslid mag een tabblad bewerken als zijn/haar functie-tekst het patroon hieronder bevat (niet hoofdlettergevoelig).
        {readOnly && ' Alleen de Voorzitter kan dit aanpassen.'}
      </p>
      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Functie bevat</th>
              {EDITEERBARE_TABS.map(t => <th key={t.id} className="px-2 py-2 font-medium text-xs text-center" style={{ color: C.inkSoft }}>{t.label}</th>)}
              <th></th>
            </tr>
          </thead>
          <tbody>
            {rolpermissies.map(r => (
              <tr key={r.id} className="border-b last:border-0" style={{ borderColor: C.border }}>
                <td className="px-3 py-2">
                  <input disabled={readOnly} className="rounded border px-2 py-1 text-sm" style={inputStyle}
                    value={r.patroon} onChange={e => updateRegel(r.id, { patroon: e.target.value })} />
                </td>
                {EDITEERBARE_TABS.map(t => (
                  <td key={t.id} className="px-2 py-2 text-center">
                    <input type="checkbox" disabled={readOnly} checked={r.tabs.includes(t.id)} onChange={() => toggleTab(r, t.id)} />
                  </td>
                ))}
                <td className="px-3 py-2 text-right">
                  {!readOnly && <button onClick={() => removeRegel(r.id)} style={{ color: C.rose }}><Trash2 size={13} /></button>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!rolpermissies.length && <EmptyState icon={Settings} text="Nog geen rollen ingesteld." />}
      </Card>
      {!readOnly && (
        showNew ? (
          <PromptModal title="Nieuwe rol" label="Functie bevat (bv. 'penningmeester')" placeholder="tekstpatroon" onSave={addRegel} onCancel={() => setShowNew(false)} />
        ) : (
          <Btn icon={Plus} tone="outline" onClick={() => setShowNew(true)}>Rol toevoegen</Btn>
        )
      )}
    </div>
  );
}
