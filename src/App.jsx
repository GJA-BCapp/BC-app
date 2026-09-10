import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, LineChart, Line, Legend } from 'recharts';
import {
  Users, PiggyBank, Wallet, TrendingUp, TrendingDown, Plus, X, Search,
  Download, Calendar, Pencil, Trash2, Baby, FileSpreadsheet, LayoutDashboard,
  Filter, ArrowUpRight, ArrowDownRight, Grid3x3, ListChecks, Landmark, Printer,
  ClipboardList, ArrowLeft, CheckSquare, Square, FileText, Link2,
  Settings, History, RotateCcw, DatabaseBackup, CheckCircle2, AlertCircle, Palette, Clock3, Upload, Phone, MessageCircle, ChevronUp, ChevronDown, Copy, Lock, Sliders,
  ExternalLink
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

const APP_VERSIE = '10-09-2026';
const SEED_SLOTS = ["ma 11.00 - 16.00","di 10.00 - 16.00","wo 09.00 - 12.30","wo 19.00 - 22.00","do 09.30 - 16.00","do 19.00 - 22.00"];
const SEED_AGENDAPUNTEN_VOORAF = ["Opening", "Mededelingen", "Vaststellen agenda", "Notulen vorige vergadering"];
const SEED_AGENDAPUNTEN_AFSLUITEND = ["Rondvraag", "Sluiting"];

/* Welke tabbladen mag iemand BEWERKEN, op basis van hun functie (Dashboard en Rapportage
   zijn altijd overal zichtbaar/exporteerbaar — dat zijn geen bewerkacties).
   Herken je hier je eigen bestuursfuncties niet in, vraag dan gewoon om deze lijst aan te passen. */
const EDITEERBARE_TABS = [
  { id: 'leden', label: 'Leden' }, { id: 'workshops', label: 'Workshops' },
  { id: 'vergaderingen', label: 'Vergaderingen' }, { id: 'financien', label: 'Financiën' },
  { id: 'begroting', label: 'Begroting' }, { id: 'kalender', label: 'Kalender' },
];
const SEED_ROLPERMISSIES = [
  { id: 1, patroon: 'voorzitter', tabs: ['leden', 'workshops', 'vergaderingen', 'financien', 'begroting', 'kalender'] },
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
  for (const offset of [0, -1, 1]) {
    if (await hotp(bytes, stap + offset) === schoon) return true;
  }
  return false;
}
function wachtwoordVoldoetAanEisen(w) {
  return !!w && w.length >= 8 && /[A-Z]/.test(w) && /[a-z]/.test(w) && /[0-9]/.test(w) && /[^A-Za-z0-9]/.test(w);
}

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
  if (typeof opgeslagen === 'string') return opgeslagen === wachtwoord;
  if (!opgeslagen || !opgeslagen.hash || !opgeslagen.salt) return false;
  const { hash } = await hashWachtwoord(wachtwoord, opgeslagen.salt);
  return hash === opgeslagen.hash;
}

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

/* =========================================================================
   KALENDER — gedeelde opbouw van de activiteitenlijst (workshops, vergaderingen
   en overige activiteiten), gebruikt door het Kalender-tabblad, het dashboard-
   widget en het losse agenda-venster. Eén plek, zodat die drie weergaven altijd
   hetzelfde tonen.
========================================================================= */
const KALENDER_TYPE_LABEL = { workshop: 'Workshop', vergadering: 'Vergadering', dagdeel: 'Schildergroep', overig: 'Overig' };
const KALENDER_TYPE_TONE = { workshop: 'ochre', vergadering: 'clay', dagdeel: 'rose', overig: 'sage' };
const DAG_AFKORTING_NAAR_NUMMER = { zo: 0, ma: 1, di: 2, wo: 3, do: 4, vr: 5, za: 6 };

/* Genereert terugkerende kalendergebeurtenissen voor de schildergroepen/dagdelen (bv. "wo 09.00 - 12.30")
   binnen een periode — dagdelen zelf zijn geen losse datums, dus die worden hier per week "uitgerekend". */
function genereerDagdeelItems(dagdelen, van, tot) {
  if (!van || !tot) return [];
  const items = [];
  const eind = new Date(tot);
  dagdelen.forEach(slot => {
    const afk = slot.trim().slice(0, 2).toLowerCase();
    const dagNummer = DAG_AFKORTING_NAAR_NUMMER[afk];
    if (dagNummer == null) return;
    let cursor = new Date(van);
    const verschil = (dagNummer - cursor.getDay() + 7) % 7;
    cursor.setDate(cursor.getDate() + verschil);
    while (cursor <= eind) {
      const datumIso = cursor.toISOString().slice(0, 10);
      items.push({
        id: `dagdeel-${slot}-${datumIso}`, type: 'dagdeel', titel: slot,
        datum: datumIso, datumTot: datumIso, locatie: '', detail: 'Schildergroep',
      });
      cursor.setDate(cursor.getDate() + 7);
    }
  });
  return items;
}
function dagdeelGenerentieGrenzen(grenzen, tijd) {
  if (grenzen.van && grenzen.tot) return { van: grenzen.van, tot: grenzen.tot };
  const vandaag = new Date();
  const iso = d => d.toISOString().slice(0, 10);
  const min90 = new Date(vandaag); min90.setDate(min90.getDate() - 90);
  const plus90 = new Date(vandaag); plus90.setDate(plus90.getDate() + 90);
  if (tijd === 'geweest') return { van: iso(min90), tot: iso(vandaag) };
  if (tijd === 'aankomend') return { van: iso(vandaag), tot: iso(plus90) };
  return { van: iso(min90), tot: iso(plus90) };
}

function bouwKalenderItems({ workshops, vergaderingen, overigeActiviteiten }) {
  const items = [];
  (workshops || []).forEach(w => {
    (w.datums || []).forEach(datum => {
      if (!datum) return;
      items.push({
        id: `workshop-${w.id}-${datum}`, type: 'workshop', titel: w.titel,
        datum, datumTot: datum, locatie: w.locatie || '',
        detail: [w.soort, w.dagdeel].filter(Boolean).join(' · '),
      });
    });
  });
  (vergaderingen || []).forEach(v => {
    if (!v.datum) return;
    items.push({
      id: `vergadering-${v.id}`, type: 'vergadering', titel: v.titel,
      datum: v.datum, datumTot: v.datum, locatie: v.locatie || '',
      detail: `${(v.agendapunten || []).length} agendapunten`,
    });
  });
  (overigeActiviteiten || []).forEach(a => {
    if (!a.datumVan) return;
    items.push({
      id: `overig-${a.id}`, type: 'overig', titel: a.titel,
      datum: a.datumVan, datumTot: a.datumTot || a.datumVan, locatie: a.locatie || '',
      detail: a.omschrijving || '', ref: a,
    });
  });
  return items.sort((a, b) => a.datum.localeCompare(b.datum) || a.titel.localeCompare(b.titel));
}

function kalenderPeriodeGrenzen(periode, vanaf, tot) {
  const vandaag = new Date();
  const iso = d => d.toISOString().slice(0, 10);
  if (periode === 'dit_jaar') return { van: `${vandaag.getFullYear()}-01-01`, tot: `${vandaag.getFullYear()}-12-31` };
  if (periode === 'komend_jaar') return { van: `${vandaag.getFullYear() + 1}-01-01`, tot: `${vandaag.getFullYear() + 1}-12-31` };
  if (periode === 'komende_3_maanden') {
    const eind = new Date(vandaag); eind.setMonth(eind.getMonth() + 3);
    return { van: iso(vandaag), tot: iso(eind) };
  }
  if (periode === 'aangepast') return { van: vanaf || null, tot: tot || null };
  return { van: null, tot: null };
}

function kalenderGroepeerPerMaand(items) {
  const groepen = [];
  let huidig = null;
  items.forEach(it => {
    const sleutel = it.datum.slice(0, 7);
    if (!huidig || huidig.sleutel !== sleutel) {
      const [jaar, maand] = sleutel.split('-');
      const naam = MONTH_NAMES[Number(maand) - 1] || '';
      huidig = { sleutel, label: `${naam.charAt(0).toUpperCase()}${naam.slice(1)} ${jaar}`, items: [] };
      groepen.push(huidig);
    }
    huidig.items.push(it);
  });
  return groepen;
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
let opslagWachtrij = Promise.resolve();
function planOpslag(taak) {
  opslagWachtrij = opslagWachtrij.then(async () => {
    try { await taak(); } catch (e) { /* taak rapporteert zijn eigen resultaat */ }
    await new Promise(r => setTimeout(r, 180));
  });
  return opslagWachtrij;
}

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
  overigeActiviteiten: 'bladels:overigeactiviteiten',
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
  const [stap, setStap] = useState('pincode');
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
  { id: 'kalender', label: 'Kalender', icon: Calendar },
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
  const [watIsNieuw, setWatIsNieuw] = useStored('bladels:watisnieuw', { tekst: '', laatstBijgewerkt: null }, true, flash);
  const [watIsNieuwGezien, setWatIsNieuwGezien] = useStored('bladels:watisnieuw-gezien', {}, true);
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
  const [overigeActiviteiten, setOverigeActiviteiten] = useStored('bladels:overigeactiviteiten', [], true, flash);
  const [actielijst, setActielijst] = useStored('bladels:actielijst', [], true, flash);
  const [contributies, setContributies] = useStored('bladels:contributies', [], true, flash);
  const [pins, setPins] = useStored('bladels:pins', {}, true, flash);
  const [beveiliging, setBeveiliging] = useStored('bladels:beveiliging', { tfaVerplicht: false, tfaVertrouwensdagen: 30, wachtwoordEisen: false }, true, flash);
  const [standaarden, setStandaarden] = useStored('bladels:standaarden', {
    idleTimeoutMinuten: 3, logoHoogteCm: 2.5,
    verenigingWebsite: 'www.BladelsCreatief.nl', verenigingEmail: 'info@bladelscreatief.nl',
    verenigingAdres: 'Roodborstje 9 Bladel', verenigingIban: 'NL13 RABO 0107 5409 08', jaarcontributie: 105,
  }, true, flash);
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

  useEffect(() => {
    if (!ingelogd) return;
    const interval = setInterval(() => stuurSessieSignaal(fullName(ingelogd), false), 45000);
    return () => clearInterval(interval);
  }, [ingelogd]);

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

  const apartVenster = typeof window !== 'undefined' && new URLSearchParams(window.location.search).get('venster') === 'agenda';
  if (apartVenster) {
    return <AgendaVenster workshops={workshops} vergaderingen={vergaderingen} overigeActiviteiten={overigeActiviteiten}
      ingelogd={ingelogd} onLogout={() => handleLogout('handmatig')} />;
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
        {watIsNieuw.tekst && watIsNieuw.laatstBijgewerkt && watIsNieuwGezien[fullName(ingelogd)] !== watIsNieuw.laatstBijgewerkt && (
          <div className="mb-4 rounded-xl border px-4 py-3 flex items-start gap-3" style={{ background: C.paperDim, borderColor: C.clay }}>
            <Sliders size={16} style={{ color: C.clay, flexShrink: 0, marginTop: 2 }} />
            <div className="flex-1">
              <p className="text-sm font-semibold mb-0.5" style={{ color: C.ink }}>Wat is er nieuw</p>
              <p className="text-sm" style={{ color: C.ink, whiteSpace: 'pre-wrap' }}>{watIsNieuw.tekst}</p>
            </div>
            <button onClick={() => setWatIsNieuwGezien({ ...watIsNieuwGezien, [fullName(ingelogd)]: watIsNieuw.laatstBijgewerkt })}
              className="p-1 rounded hover:bg-black/5 flex-shrink-0" style={{ color: C.inkSoft }}><X size={16} /></button>
          </div>
        )}
        {tab === 'dashboard' && <Dashboard members={members} workshops={workshops} inschrijvingen={inschrijvingen} tx={tx} budget={budget} setTab={setTab} dagdelen={dagdelen}
          vergaderingen={vergaderingen} overigeActiviteiten={overigeActiviteiten} />}
        {tab === 'leden' && <LedenTab members={members} setMembers={setMembers} readOnly={readOnly('leden')}
          contributies={contributies} setContributies={setContributies} boekjaren={boekjaren} tx={tx} setTx={setTx} accounts={accounts}
          dagdelen={dagdelen} standaarden={standaarden} ingelogd={ingelogd}
          initialQuery={pendingQuery} onTrash={trashIt} onLog={logAction} />}
        {tab === 'workshops' && <WorkshopsTab members={members} workshops={workshops} setWorkshops={setWorkshops}
          inschrijvingen={inschrijvingen} setInschrijvingen={setInschrijvingen} tx={tx} setTx={setTx} accounts={accounts}
          workshopSoorten={workshopSoorten} logoHoogteCm={standaarden.logoHoogteCm}
          workshopSortering={workshopSortering} setWorkshopSortering={setWorkshopSortering}
          readOnly={readOnly('workshops')} initialQuery={pendingQuery} onTrash={trashIt} onLog={logAction} />}
        {tab === 'vergaderingen' && <VergaderingenTab members={members} vergaderingen={vergaderingen} setVergaderingen={setVergaderingen} actielijst={actielijst} setActielijst={setActielijst}
          agendapuntenVooraf={agendapuntenVooraf} agendapuntenAfsluitend={agendapuntenAfsluitend} logoHoogteCm={standaarden.logoHoogteCm}
          readOnly={readOnly('vergaderingen')} onTrash={trashIt} onLog={logAction} />}
        {tab === 'kalender' && <KalenderTab workshops={workshops} vergaderingen={vergaderingen}
          overigeActiviteiten={overigeActiviteiten} setOverigeActiviteiten={setOverigeActiviteiten} dagdelen={dagdelen}
          readOnly={readOnly('kalender')} onTrash={trashIt} onLog={logAction} />}
        {tab === 'financien' && <FinancienTab tx={tx} setTx={setTx} accounts={accounts} boekjaren={boekjaren} setBoekjaren={setBoekjaren}
          members={members} contributies={contributies} setContributies={setContributies}
          workshops={workshops} inschrijvingen={inschrijvingen} setInschrijvingen={setInschrijvingen}
          readOnly={readOnly('financien')} initialQuery={pendingQuery} onTrash={trashIt} onLog={logAction} />}
        {tab === 'begroting' && <BegrotingTab budget={budget} setBudget={setBudget} tx={tx} boekjaren={boekjaren} setBoekjaren={setBoekjaren} begrotingKoppelingen={begrotingKoppelingen} readOnly={readOnly('begroting')} onLog={logAction} />}
        {tab === 'rapportage' && <RapportageTab tx={tx} accounts={accounts} members={members} workshops={workshops} inschrijvingen={inschrijvingen} budget={budget} begrotingKoppelingen={begrotingKoppelingen} logoHoogteCm={standaarden.logoHoogteCm} />}
        {tab === 'instellingen' && (
          <InstellingenTab isVoorzitter={isVoorzitter} ingelogd={ingelogd} rolpermissies={rolpermissies} setRolpermissies={setRolpermissies}
            beveiliging={beveiliging} setBeveiliging={setBeveiliging} standaarden={standaarden} setStandaarden={setStandaarden}
            watIsNieuw={watIsNieuw} setWatIsNieuw={setWatIsNieuw}
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
            overigeActiviteiten={overigeActiviteiten} setOverigeActiviteiten={setOverigeActiviteiten}
            magLeden={magBewerken.has('leden')} magWorkshops={magBewerken.has('workshops')} magFinancien={magBewerken.has('financien')} magVergaderingen={magBewerken.has('vergaderingen')}
            magLedenImporteren={magBewerken.has('leden')} magWorkshopsImporteren={magBewerken.has('workshops')} magFinancienImporteren={magBewerken.has('financien')}
            backupData={{ members, workshops, inschrijvingen, tx, accounts, budget, boekjaren, vergaderingen, actielijst, contributies, rolpermissies, begrotingKoppelingen, overigeActiviteiten }}
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
function Dashboard({ members, workshops, inschrijvingen, tx, budget, setTab, dagdelen, vergaderingen, overigeActiviteiten }) {
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

  const vandaagIso = new Date().toISOString().slice(0, 10);
  const komendeActiviteiten = bouwKalenderItems({ workshops, vergaderingen, overigeActiviteiten })
    .filter(it => it.datum >= vandaagIso).slice(0, 5);

  function openAgendaVenster() {
    window.open(`${window.location.origin}${window.location.pathname}?venster=agenda`, 'BladelsCreatiefAgenda', 'width=380,height=640,resizable=yes,scrollbars=yes');
  }

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

      <Card className="p-4">
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-2">
            <Calendar size={16} style={{ color: C.clay }} />
            <h3 className="font-semibold" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Komende activiteiten</h3>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={openAgendaVenster} className="text-xs underline flex items-center gap-1" style={{ color: C.clay }}>
              <ExternalLink size={12} /> Apart venster
            </button>
            <button onClick={() => setTab('kalender')} className="text-xs underline" style={{ color: C.clay }}>Volledige kalender →</button>
          </div>
        </div>
        {komendeActiviteiten.length ? (
          <div className="grid sm:grid-cols-2 lg:grid-cols-5 gap-2">
            {komendeActiviteiten.map(it => (
              <div key={it.id} className="px-3 py-2 rounded-lg" style={{ background: C.paperDim }}>
                <Badge tone={KALENDER_TYPE_TONE[it.type]}>{KALENDER_TYPE_LABEL[it.type]}</Badge>
                <p className="text-sm font-medium mt-1 truncate" title={it.titel}>{it.titel}</p>
                <p className="text-xs" style={{ color: C.inkSoft }}>{fmtDate(it.datum)}{it.locatie ? ' · ' + it.locatie : ''}</p>
              </div>
            ))}
          </div>
        ) : (
          <p className="text-sm" style={{ color: C.inkSoft }}>Geen aankomende activiteiten gevonden.</p>
        )}
      </Card>
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
function LedenTab({ members, setMembers, readOnly, contributies, setContributies, boekjaren, tx, setTx, accounts, dagdelen, standaarden, ingelogd, initialQuery, onTrash, onLog }) {
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

  const [toonWelkomstbrief, setToonWelkomstbrief] = useState(null);
  const [delId, setDelId] = useState(null);
  function removeMember(id) {
    const m = members.find(x => x.id === id);
    setMembers(members.filter(x => x.id !== id));
    if (m) { onTrash('lid', m); onLog(`Lid verwijderd: ${fullName(m)}`, 'leden'); }
  }
  const [toonUitschrijfbrief, setToonUitschrijfbrief] = useState(null);
  function saveMember(data) {
    const wasNieuw = !data.id;
    const werdInactief = editing && editing.status !== 'inactief' && data.status === 'inactief';
    let opgeslagen = data;
    if (data.id) {
      setMembers(members.map(m => m.id === data.id ? data : m));
      onLog(`Lid bewerkt: ${fullName(data)}`, 'leden');
    } else {
      opgeslagen = { ...data, id: uid(members) };
      setMembers([...members, opgeslagen]);
      onLog(`Lid toegevoegd: ${fullName(data)}`, 'leden');
    }
    setShowForm(false); setEditing(null);
    if (wasNieuw) setToonWelkomstbrief(opgeslagen);
    else if (werdInactief) setToonUitschrijfbrief(opgeslagen);
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
                      {!readOnly && m.status === 'inactief' && (
                        <button onClick={() => setDelId(m.id)} title="Verwijderen (alleen mogelijk bij status Inactief)" className="p-1.5 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={14} /></button>
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
      {toonWelkomstbrief && (
        <WelkomstbriefModal lid={toonWelkomstbrief} standaarden={standaarden} ingelogd={ingelogd} onClose={() => setToonWelkomstbrief(null)} />
      )}
      {toonUitschrijfbrief && (
        <UitschrijfbriefModal lid={toonUitschrijfbrief} standaarden={standaarden} ingelogd={ingelogd} onClose={() => setToonUitschrijfbrief(null)} />
      )}
      {delId != null && (
        <ConfirmModal message="Dit inactieve lid definitief verwijderen? Het lid is terug te vinden in de prullenbak (Instellingen)." onConfirm={() => { removeMember(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
    </div>
  );
}

function WelkomstbriefModal({ lid, standaarden, ingelogd, onClose }) {
  const [dagdeelKeuze, setDagdeelKeuze] = useState((lid.dagdelen || [])[0] || '');
  const [ingangsdatum, setIngangsdatum] = useState(lid.lidsinds || new Date().toISOString().slice(0, 10));
  const [afzenderNaam, setAfzenderNaam] = useState(ingelogd ? fullName(ingelogd) : '');
  const [afzenderFunctie, setAfzenderFunctie] = useState(ingelogd ? (ingelogd.functie || '') : '');
  function download() {
    const body = welkomstbriefWordHtml(lid, standaarden, afzenderNaam, afzenderFunctie, dagdeelKeuze, ingangsdatum, standaarden.logoHoogteCm);
    downloadWordDoc({ titel: 'Welkomstbrief', filename: `BladelsCreatief_Welkomstbrief_${fullName(lid).replace(/\s+/g, '_')}.doc`, bodyHtml: body, footerHtml: ledenbriefFooterHtml(standaarden), margeCm: 0.75 });
    onClose();
  }
  return (
    <Modal title={`Welkomstbrief voor ${fullName(lid)}`} onClose={onClose}>
      <p className="text-sm mb-3" style={{ color: C.inkSoft }}>Nieuw lid toegevoegd — wil je meteen een welkomstbrief met de ingevoerde gegevens downloaden?</p>
      <div className="space-y-3">
        <Field label="Dagdeel/groep (voor de tekst 'veel plezier bij de ... groep')">
          <input className={inputCls} style={inputStyle} value={dagdeelKeuze} onChange={e => setDagdeelKeuze(e.target.value)} />
        </Field>
        <Field label="Ingangsdatum contributie"><input type="date" className={inputCls} style={inputStyle} value={ingangsdatum} onChange={e => setIngangsdatum(e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Afzender naam"><input className={inputCls} style={inputStyle} value={afzenderNaam} onChange={e => setAfzenderNaam(e.target.value)} /></Field>
          <Field label="Afzender functie"><input className={inputCls} style={inputStyle} value={afzenderFunctie} onChange={e => setAfzenderFunctie(e.target.value)} /></Field>
        </div>
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Overslaan</Btn>
        <Btn tone="sage" icon={FileText} onClick={download}>Downloaden</Btn>
      </div>
    </Modal>
  );
}

function UitschrijfbriefModal({ lid, standaarden, ingelogd, onClose }) {
  const [einddatum, setEinddatum] = useState(lid.eindelidmaat || new Date().toISOString().slice(0, 10));
  const [afzenderNaam, setAfzenderNaam] = useState(ingelogd ? fullName(ingelogd) : '');
  function download() {
    const body = uitschrijfbriefWordHtml(lid, einddatum, afzenderNaam, standaarden.logoHoogteCm);
    downloadWordDoc({ titel: 'Uitschrijfbevestiging', filename: `BladelsCreatief_Uitschrijfbevestiging_${fullName(lid).replace(/\s+/g, '_')}.doc`, bodyHtml: body, footerHtml: ledenbriefFooterHtml(standaarden), margeCm: 0.75 });
    onClose();
  }
  return (
    <Modal title={`Uitschrijfbevestiging voor ${fullName(lid)}`} onClose={onClose}>
      <p className="text-sm mb-3" style={{ color: C.inkSoft }}>Dit lid is op inactief gezet — wil je meteen een uitschrijfbevestiging downloaden?</p>
      <div className="space-y-3">
        <Field label="Einddatum lidmaatschap"><input type="date" className={inputCls} style={inputStyle} value={einddatum} onChange={e => setEinddatum(e.target.value)} /></Field>
        <Field label="Afzender naam"><input className={inputCls} style={inputStyle} value={afzenderNaam} onChange={e => setAfzenderNaam(e.target.value)} /></Field>
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Overslaan</Btn>
        <Btn tone="sage" icon={FileText} onClick={download}>Downloaden</Btn>
      </div>
    </Modal>
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
        <Field label="Geboortedatum"><input type="date" className={inputCls} style={inputStyle} value={f.gebdatum || ''} onChange={e => upd('gebdatum', e.target.value)} /></Field>
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

function logoImgTag(hoogteCm) {
  const h = Math.round((hoogteCm / 2.54) * 96);
  const w = Math.round(h * (400 / 208));
  return `<img src="${LOGO_URI}" width="${w}" height="${h}" style="height:${hoogteCm}cm;max-height:${hoogteCm}cm;width:${w}px;display:block;margin-bottom:4px;" />`;
}
function downloadWordDoc({ titel, filename, bodyHtml, footerHtml, margeCm }) {
  const logoBase64 = LOGO_URI.split(',')[1] || '';
  const bodyMetCid = bodyHtml.split(LOGO_URI).join('cid:logo.png');
  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(titel)}</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
  @page { margin: ${margeCm || 2.2}cm; }
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
  .brief-voettekst { border-top: 1px solid #999999; padding-top: 6px; margin-top: 0; font-size: 10pt; }
  .brief p { margin: 0 0 2px; line-height: 1.35; }
  .brief .witregel { margin: 0; line-height: 1.35; }
  .brief table td { border: none; padding: 0 6px 0 0; line-height: 1.35; }
  .brief table { margin: 0; }
</style>
</head>
<body>
${bodyMetCid}
${footerHtml ? `<div class="brief-voettekst">${footerHtml}</div>` : ''}
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
    const a = document.createElement('a');
    a.href = url;
    a.download = `BladelsCreatief_${titel.replace(/\s+/g, '_')}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
  setTimeout(() => URL.revokeObjectURL(url), 8000);
}

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

/* Berekent een voorstel voor de resterende jaarcontributie op basis van de maand van
   binnenkomst (aantal resterende maanden incl. de maand zelf, gedeeld door 12) — komt overeen
   met hoe BladelsCreatief dit in de praktijk al berekent (bv. instap 1 okt = 3/12 van het jaarbedrag). */
function berekenProrataContributie(jaarbedrag, vanafDatumIso) {
  if (!vanafDatumIso) return jaarbedrag;
  const maand = new Date(vanafDatumIso).getMonth() + 1;
  const resterendeMaanden = 12 - maand + 1;
  return Math.round((jaarbedrag / 12) * resterendeMaanden * 100) / 100;
}

function ledenbriefFooterHtml(standaarden) {
  return `<table style="margin-top:18px;">
<tr><td style="border:none;width:50%;">Website: <a href="https://${standaarden.verenigingWebsite.replace(/^https?:\/\//, '')}">${escapeHtml(standaarden.verenigingWebsite)}</a></td>
<td style="border:none;text-align:right;">Adres: ${escapeHtml(standaarden.verenigingAdres)}</td></tr>
<tr><td style="border:none;">Mail: <a href="mailto:${standaarden.verenigingEmail}">${escapeHtml(standaarden.verenigingEmail)}</a></td>
<td style="border:none;text-align:right;">Rabobank ${escapeHtml(standaarden.verenigingIban)}</td></tr>
</table>`;
}

function welkomstbriefWordHtml(lid, standaarden, afzenderNaam, afzenderFunctie, dagdeelKeuze, ingangsdatum, logoHoogteCm) {
  const vandaag = new Date().toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' });
  const contributieResterend = berekenProrataContributie(standaarden.jaarcontributie, ingangsdatum || lid.lidsinds);
  const jaar = new Date().getFullYear();
  const wr = '<p class="witregel">&nbsp;</p>';
  return `${logoImgTag(logoHoogteCm)}
<div class="brief">
${wr}
<p>Bladel, ${vandaag}</p>
${wr}${wr}
<p>Hallo ${escapeHtml(lid.voornaam)},</p>
${wr}
<p>Welkom bij BladelsCreatief!</p>
<p>Bedankt voor je inschrijving, ik heb het formulier ontvangen — deze brief is ter controle van je gegevens.</p>
<table><tr><td>${escapeHtml(fullName(lid))}</td><td>${escapeHtml(lid.email || '—')}</td></tr>
<tr><td colspan="2">${escapeHtml(lid.adres || '—')}</td></tr>
<tr><td>${escapeHtml(lid.postcode || '')} ${escapeHtml(lid.woonplaats || '')}</td><td>${escapeHtml(lid.telefoon || '—')} · ${lid.gebdatum ? fmtDate(lid.gebdatum) : '—'}</td></tr></table>
<p>Graag een oké retour als de gegevens juist zijn, of het ontbrekende aanvullen.</p>
${wr}
<p>We vinden het fijn je bij BladelsCreatief te begroeten. Er ligt een roze map in de kast in het atelier, daarin vind je o.a. de ledenlijst, huishoudelijk reglement en andere zaken ter inzage. Kijk hiervoor ook naar onze website.</p>
${wr}
<p>Je mailadres wordt alleen gebruikt voor brieven en mededelingen betreffende BladelsCreatief. Met je mobiele nummer voegen we je toe aan de groeps-app voor de leden van BladelsCreatief. Dit is een één-richting groeps-app die door het bestuur gebruikt wordt om mededelingen te doen of een peiling uit te zetten. Je kunt er niet op reageren en zult dus ook niet 'overladen worden' met reacties van andere leden. Reclame en andere zaken worden níet per mail doorgestuurd, deze komen in de roze map op de tafel.</p>
${wr}
<p>Veel plezier bij de ${escapeHtml(dagdeelKeuze || (lid.dagdelen || []).join(', ') || '(dagdeel)')} groep, veel inspiratie en succes toegewenst!</p>
${wr}
<p>De contributie wordt per jaar voldaan en bedraagt voor ${jaar} ${euro(standaarden.jaarcontributie)}. Ik wil je vragen om <strong>${euro(contributieResterend)}</strong> als contributiebedrag voor de rest van dit jaar (vanaf ${fmtDate(ingangsdatum || lid.lidsinds)}) op onderstaand bankrekeningnummer te voldoen.</p>
${wr}
<p>Als lid van BladelsCreatief kun je in principe op alle tijden dat we van het atelier gebruikmaken deelnemen aan de bestaande groepen. Het is gewenst om even te overleggen als je van groep wilt wisselen.</p>
${wr}
<p>Met vriendelijke groet,</p>
${wr}
<p>${escapeHtml(afzenderNaam)}<br/>${escapeHtml(afzenderFunctie)} BladelsCreatief</p>
</div>`;
}

function uitschrijfbriefWordHtml(lid, einddatum, afzenderNaam, logoHoogteCm) {
  const vandaag = new Date().toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' });
  const wr = '<p class="witregel">&nbsp;</p>';
  return `${logoImgTag(logoHoogteCm)}
<div class="brief">
${wr}
<p>Bladel, ${vandaag}</p>
${wr}${wr}
<p>Hallo ${escapeHtml(lid.voornaam)},</p>
${wr}
<p>De opzegging van je lidmaatschap heb ik ontvangen. Bedankt voor de tijd die je bij BladelsCreatief geschilderd hebt. Ik hoop dat je ervan genoten hebt. Wellicht vind je een passend alternatief waarbij je je hobby met veel plezier kunt uitoefenen.</p>
${wr}
<p>Ik zal je lidmaatschap per ${fmtDate(einddatum)} beëindigen. Tot die tijd blijf je lid en kun je blijven schilderen en deelnemen aan andere activiteiten van onze vereniging.</p>
${wr}
<p>Met vriendelijke groet,</p>
<p>Namens het bestuur van BladelsCreatief</p>
${wr}
<p>${escapeHtml(afzenderNaam)}</p>
</div>`;
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
  const [tijd, setTijd] = useState('aankomend');
  const openActies = actielijst.filter(a => a.status !== 'klaar');
  const vandaagIso = new Date().toISOString().slice(0, 10);
  let gefilterd;
  if (tijd === 'recent') {
    const laatsteAfgelopen = vergaderingen
      .filter(v => (v.datum || '') < vandaagIso)
      .sort((a, b) => (b.datum || '').localeCompare(a.datum || ''))
      .slice(0, 2);
    const komend = vergaderingen.filter(v => (v.datum || '') >= vandaagIso);
    gefilterd = [...laatsteAfgelopen, ...komend];
  } else {
    gefilterd = vergaderingen.filter(v => {
      if (tijd === 'aankomend') return (v.datum || '') >= vandaagIso;
      if (tijd === 'geweest') return (v.datum || '') < vandaagIso;
      return true;
    });
  }
  const sorted = [...gefilterd].sort((a, b) => (tijd === 'aankomend' || tijd === 'recent')
    ? (a.datum || '').localeCompare(b.datum || '')
    : (b.datum || '').localeCompare(a.datum || ''));

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
        <div className="flex flex-wrap gap-2 items-center">
          <select value={tijd} onChange={e => setTijd(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="aankomend">Aankomend</option>
            <option value="recent">Recent (laatste 2) + aankomend</option>
            <option value="geweest">Geweest</option>
            <option value="alle">Alle</option>
          </select>
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
            <EmptyState icon={ClipboardList} text={
              tijd === 'aankomend' ? 'Geen aankomende vergaderingen gepland.'
                : tijd === 'geweest' ? 'Nog geen eerdere vergaderingen gevonden.'
                : 'Nog geen vergaderingen vastgelegd. Maak een nieuwe vergadering aan om agenda en notulen te koppelen.'
            } />
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
   KALENDER-TABBLAD — gecombineerd overzicht van workshops, vergaderingen en
   overige activiteiten (nieuwjaarsborrel e.d.), met filters, export naar
   Word/PDF en een los te openen agenda-venster.
========================================================================= */
function KalenderTab({ workshops, vergaderingen, overigeActiviteiten, setOverigeActiviteiten, dagdelen, readOnly, onTrash, onLog }) {
  const [typeFilter, setTypeFilter] = useState({ workshop: true, vergadering: true, dagdeel: false, overig: true });
  const [periode, setPeriode] = useState('dit_jaar');
  const [vanaf, setVanaf] = useState(new Date().toISOString().slice(0, 10));
  const [tot, setTot] = useState(new Date().toISOString().slice(0, 10));
  const [tijd, setTijd] = useState('aankomend');
  const [showNew, setShowNew] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [delId, setDelId] = useState(null);

  const vandaagIso = new Date().toISOString().slice(0, 10);
  const grenzen = kalenderPeriodeGrenzen(periode, vanaf, tot);
  const alleItems = bouwKalenderItems({ workshops, vergaderingen, overigeActiviteiten });
  const dagdeelGrenzen = dagdeelGenerentieGrenzen(grenzen, tijd);
  const dagdeelItems = typeFilter.dagdeel ? genereerDagdeelItems(dagdelen, dagdeelGrenzen.van, dagdeelGrenzen.tot) : [];
  const alleItemsMetDagdelen = [...alleItems, ...dagdeelItems].sort((a, b) => a.datum.localeCompare(b.datum) || a.titel.localeCompare(b.titel));

  const gefilterd = alleItemsMetDagdelen.filter(it => {
    if (!typeFilter[it.type]) return false;
    if (tijd === 'aankomend' && it.datum < vandaagIso) return false;
    if (tijd === 'geweest' && it.datum >= vandaagIso) return false;
    if (grenzen.van && it.datum < grenzen.van) return false;
    if (grenzen.tot && it.datum > grenzen.tot) return false;
    return true;
  });
  const groepen = kalenderGroepeerPerMaand(gefilterd);

  function toggleType(t) { setTypeFilter(f => ({ ...f, [t]: !f[t] })); }

  function saveActiviteit(f) {
    if (editItem) {
      setOverigeActiviteiten(overigeActiviteiten.map(a => a.id === editItem.id ? { ...a, ...f } : a));
      onLog(`Activiteit bijgewerkt: ${f.titel}`, 'kalender');
    } else {
      const id = uid(overigeActiviteiten);
      setOverigeActiviteiten([...overigeActiviteiten, { id, ...f }]);
      onLog(`Activiteit toegevoegd: ${f.titel} (${f.datumVan})`, 'kalender');
    }
    setShowNew(false);
    setEditItem(null);
  }
  function removeActiviteit(id) {
    const a = overigeActiviteiten.find(x => x.id === id);
    setOverigeActiviteiten(overigeActiviteiten.filter(x => x.id !== id));
    if (a) { onTrash('activiteit', a); onLog(`Activiteit verwijderd: ${a.titel}`, 'kalender'); }
  }

  function exporteerWord() {
    const rijenHtml = gefilterd.map(it => `<tr><td>${fmtDate(it.datum)}${it.datumTot && it.datumTot !== it.datum ? ' t/m ' + fmtDate(it.datumTot) : ''}</td><td>${escapeHtml(KALENDER_TYPE_LABEL[it.type])}</td><td>${escapeHtml(it.titel)}</td><td>${escapeHtml(it.locatie || '—')}</td><td>${escapeHtml(it.detail || '—')}</td></tr>`).join('');
    const bodyHtml = `<h1>Activiteitenkalender</h1>
<p class="meta">Gegenereerd op ${new Date().toLocaleDateString('nl-NL')} · ${gefilterd.length} activiteit(en)</p>
<table>
  <thead><tr><th>Datum</th><th>Type</th><th>Titel</th><th>Locatie</th><th>Details</th></tr></thead>
  <tbody>${rijenHtml || '<tr><td colspan="5"><em>Geen activiteiten gevonden binnen deze filters.</em></td></tr>'}</tbody>
</table>`;
    downloadWordDoc({ titel: 'Activiteitenkalender', filename: `BladelsCreatief_Activiteitenkalender_${new Date().toISOString().slice(0, 10)}.doc`, bodyHtml });
  }
  function exporteerPdf() {
    downloadPrintableHtml({
      titel: 'Activiteitenkalender',
      subtitel: `${gefilterd.length} activiteit(en)`,
      header: ['Datum', 'Type', 'Titel', 'Locatie', 'Details'],
      rows: gefilterd.map(it => [
        fmtDate(it.datum) + (it.datumTot && it.datumTot !== it.datum ? ' t/m ' + fmtDate(it.datumTot) : ''),
        KALENDER_TYPE_LABEL[it.type], it.titel, it.locatie || '—', it.detail || '—',
      ]),
    });
  }
  function openAgendaVenster() {
    window.open(`${window.location.origin}${window.location.pathname}?venster=agenda`, 'BladelsCreatiefAgenda', 'width=380,height=640,resizable=yes,scrollbars=yes');
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <h2 className="font-semibold" style={{ fontFamily: 'Fraunces, serif', fontSize: 18, color: C.ink }}>Kalender</h2>
          <p className="text-xs" style={{ color: C.inkSoft }}>Workshops, vergaderingen en overige activiteiten in één overzicht.</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <Btn tone="outline" icon={ExternalLink} onClick={openAgendaVenster}>Apart venster</Btn>
          <Btn tone="outline" icon={FileText} onClick={exporteerWord} disabled={!gefilterd.length}>Word</Btn>
          <Btn tone="outline" icon={Printer} onClick={exporteerPdf} disabled={!gefilterd.length}>PDF</Btn>
          {!readOnly && <Btn icon={Plus} onClick={() => { setEditItem(null); setShowNew(true); }}>Activiteit toevoegen</Btn>}
        </div>
      </div>

      <Card className="p-3">
        <div className="flex flex-wrap items-center gap-3">
          <div className="flex gap-1.5">
            {Object.keys(KALENDER_TYPE_LABEL).map(t => (
              <button key={t} onClick={() => toggleType(t)} className="px-2.5 py-1 rounded-full text-xs border flex items-center gap-1"
                style={{ background: typeFilter[t] ? C.clay : 'transparent', color: typeFilter[t] ? '#fff' : C.ink, borderColor: typeFilter[t] ? C.clay : C.border }}>
                {typeFilter[t] ? <CheckSquare size={12} /> : <Square size={12} />} {KALENDER_TYPE_LABEL[t]}
              </button>
            ))}
          </div>
          <select value={tijd} onChange={e => setTijd(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="aankomend">Aankomend</option>
            <option value="geweest">Geweest</option>
            <option value="alle">Alle tijden</option>
          </select>
          <select value={periode} onChange={e => setPeriode(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Alle perioden</option>
            <option value="dit_jaar">Dit kalenderjaar</option>
            <option value="komend_jaar">Volgend kalenderjaar</option>
            <option value="komende_3_maanden">Komende 3 maanden</option>
            <option value="aangepast">Aangepaste periode…</option>
          </select>
          {periode === 'aangepast' && (
            <div className="flex items-center gap-1.5">
              <input type="date" className={inputCls} style={inputStyle} value={vanaf} onChange={e => setVanaf(e.target.value)} />
              <span className="text-xs" style={{ color: C.inkSoft }}>t/m</span>
              <input type="date" className={inputCls} style={inputStyle} value={tot} onChange={e => setTot(e.target.value)} />
            </div>
          )}
        </div>
      </Card>

      <div className="space-y-5">
        {groepen.map(groep => (
          <div key={groep.sleutel}>
            <h3 className="text-sm font-semibold mb-2" style={{ color: C.inkSoft }}>{groep.label}</h3>
            <div className="space-y-2">
              {groep.items.map(it => (
                <Card key={it.id} className="p-3 flex items-center justify-between gap-3">
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="w-14 text-center flex-shrink-0">
                      <p className="text-lg font-semibold leading-none" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>{it.datum.slice(8, 10)}</p>
                      <p className="text-xs" style={{ color: C.inkSoft }}>{new Date(it.datum).toLocaleDateString('nl-NL', { weekday: 'short' })}</p>
                    </div>
                    <div className="min-w-0">
                      <div className="flex items-center gap-2">
                        <Badge tone={KALENDER_TYPE_TONE[it.type]}>{KALENDER_TYPE_LABEL[it.type]}</Badge>
                        <p className="font-medium truncate" style={{ color: C.ink }}>{it.titel}</p>
                      </div>
                      <p className="text-xs truncate" style={{ color: C.inkSoft }}>{[it.locatie, it.detail].filter(Boolean).join(' · ') || '—'}</p>
                    </div>
                  </div>
                  {it.type === 'overig' && !readOnly && (
                    <div className="flex gap-1 flex-shrink-0">
                      <button onClick={() => { setEditItem(it.ref); setShowNew(true); }} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={14} /></button>
                      <button onClick={() => setDelId(it.ref.id)} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={14} /></button>
                    </div>
                  )}
                </Card>
              ))}
            </div>
          </div>
        ))}
        {!groepen.length && <EmptyState icon={Calendar} text="Geen activiteiten gevonden binnen deze filters." />}
      </div>

      {showNew && (
        <ActiviteitModal item={editItem} onSave={saveActiviteit} onClose={() => { setShowNew(false); setEditItem(null); }} />
      )}
      {delId != null && (
        <ConfirmModal message="Deze activiteit verwijderen?" onConfirm={() => { removeActiviteit(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
    </div>
  );
}

function ActiviteitModal({ item, onSave, onClose }) {
  const [f, setF] = useState(item ? {
    titel: item.titel, datumVan: item.datumVan, datumTot: item.datumTot || item.datumVan,
    locatie: item.locatie || '', omschrijving: item.omschrijving || '',
  } : {
    titel: '', datumVan: new Date().toISOString().slice(0, 10), datumTot: new Date().toISOString().slice(0, 10),
    locatie: '', omschrijving: '',
  });
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  return (
    <Modal title={item ? 'Activiteit bewerken' : 'Activiteit toevoegen'} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Titel"><input className={inputCls} style={inputStyle} placeholder="bv. Nieuwjaarsborrel" value={f.titel} onChange={e => upd('titel', e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Datum vanaf"><input type="date" className={inputCls} style={inputStyle} value={f.datumVan} onChange={e => upd('datumVan', e.target.value)} /></Field>
          <Field label="Datum tot (bij meerdaags)"><input type="date" className={inputCls} style={inputStyle} value={f.datumTot} onChange={e => upd('datumTot', e.target.value)} /></Field>
        </div>
        <Field label="Locatie"><input className={inputCls} style={inputStyle} value={f.locatie} onChange={e => upd('locatie', e.target.value)} /></Field>
        <Field label="Omschrijving (optioneel)"><textarea rows={2} className={inputCls} style={inputStyle} value={f.omschrijving} onChange={e => upd('omschrijving', e.target.value)} /></Field>
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(f)} disabled={!f.titel || !f.datumVan}>{item ? 'Opslaan' : 'Toevoegen'}</Btn>
      </div>
    </Modal>
  );
}

function AgendaVenster({ workshops, vergaderingen, overigeActiviteiten, ingelogd, onLogout }) {
  const vandaagIso = new Date().toISOString().slice(0, 10);
  const items = bouwKalenderItems({ workshops, vergaderingen, overigeActiviteiten }).filter(it => it.datum >= vandaagIso).slice(0, 25);
  const groepen = kalenderGroepeerPerMaand(items);
  return (
    <div className="min-h-screen p-3" style={{ background: C.paper, fontFamily: 'Inter, ui-sans-serif, system-ui' }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');`}</style>
      <div className="flex items-center gap-2 mb-3">
        <img src={LOGO_URI} alt="BladelsCreatief" style={{ height: 30 }} />
        <h1 className="font-semibold text-sm" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Agenda</h1>
      </div>
      <div className="space-y-4">
        {groepen.map(groep => (
          <div key={groep.sleutel}>
            <h3 className="text-xs font-semibold mb-1.5" style={{ color: C.inkSoft }}>{groep.label}</h3>
            <div className="space-y-1.5">
              {groep.items.map(it => (
                <div key={it.id} className="rounded-lg border p-2" style={{ borderColor: C.border, background: C.card }}>
                  <div className="flex items-center gap-1.5">
                    <Badge tone={KALENDER_TYPE_TONE[it.type]}>{KALENDER_TYPE_LABEL[it.type]}</Badge>
                    <span className="text-xs" style={{ color: C.inkSoft }}>{fmtDate(it.datum)}</span>
                  </div>
                  <p className="text-sm font-medium mt-0.5">{it.titel}</p>
                  {(it.locatie || it.detail) && <p className="text-xs" style={{ color: C.inkSoft }}>{[it.locatie, it.detail].filter(Boolean).join(' · ')}</p>}
                </div>
              ))}
            </div>
          </div>
        ))}
        {!groepen.length && <p className="text-xs" style={{ color: C.inkSoft }}>Geen aankomende activiteiten.</p>}
      </div>
    </div>
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
  const [q, setQ] = useState(initialQuery || '');
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [delId, setDelId] = useState(null);
  const [showNewYear, setShowNewYear] = useState(false);
  const [sortField, setSortField] = useState('datum');
  const [sortDir, setSortDir] = useState('desc');
  const [kolomFilter, setKolomFilter] = useState({ datum: '', grootboek: '', bedrag: '' });

  const filtered = tx.filter(t => {
    if (rekening !== 'alle' && t.rekening !== rekening) return false;
    if (jaar !== 'alle' && t.jaar !== Number(jaar)) return false;
    if (q && !`${t.omschrijving || ''} ${t.grootboek_naam || ''}`.toLowerCase().includes(q.toLowerCase())) return false;
    if (kolomFilter.datum && !(t.datum || '').includes(kolomFilter.datum)) return false;
    if (kolomFilter.grootboek && !`${t.grootboek_code} ${t.grootboek_naam}`.toLowerCase().includes(kolomFilter.grootboek.toLowerCase())) return false;
    if (kolomFilter.bedrag && !String(t.bedrag).includes(kolomFilter.bedrag)) return false;
    return true;
  });
  function toggleSort(field) {
    if (sortField === field) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setSortField(field); setSortDir(field === 'datum' ? 'desc' : 'asc'); }
  }
  function sortIcon(field) { return sortField !== field ? '' : (sortDir === 'asc' ? ' \u25B2' : ' \u25BC'); }
  const sorted = [...filtered].sort((a, b) => {
    let cmp;
    if (sortField === 'datum') cmp = (a.datum || '').localeCompare(b.datum || '');
    else if (sortField === 'rekening') cmp = (a.rekening || '').localeCompare(b.rekening || '');
    else if (sortField === 'grootboek') cmp = gbLabel(a).localeCompare(gbLabel(b));
    else if (sortField === 'omschrijving') cmp = (a.omschrijving || '').localeCompare(b.omschrijving || '');
    else if (sortField === 'bedrag') cmp = a.bedrag - b.bedrag;
    return sortDir === 'asc' ? cmp : -cmp;
  });
  const totaal = filtered.reduce((s, t) => s + t.bedrag, 0);

  function addYear(newYear) {
    if (!boekjaren.includes(newYear)) setBoekjaren([...boekjaren, newYear].sort((a, b) => a - b));
    setJaar(newYear);
    setShowNewYear(false);
  }
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

  const saldo = jaar !== 'alle' ? saldoBoekjaar(tx, rekening, Number(jaar)) : null;

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2">
          <select value={rekening} onChange={e => setRekening(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="908">.908 lopend</option>
            <option value="319">.319 spaar</option>
            <option value="alle">Beide rekeningen</option>
          </select>
          <select value={jaar} onChange={e => setJaar(e.target.value === 'alle' ? 'alle' : Number(e.target.value))} className={inputCls} style={inputStyle}>
            {years.map(j => <option key={j} value={j}>{j}</option>)}
            <option value="alle">Alle jaren</option>
          </select>
          {!readOnly && <button onClick={() => setShowNewYear(true)} className="text-xs underline" style={{ color: C.clay }}>+ nieuw boekjaar</button>}
          <div className="relative">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: C.inkSoft }} />
            <input value={q} onChange={e => setQ(e.target.value)} placeholder="Zoek boeking…" className={`${inputCls} pl-8`} style={{ ...inputStyle, width: 190 }} />
          </div>
        </div>
        {!readOnly && <Btn icon={Plus} onClick={() => { setEditing(null); setShowForm(true); }}>Nieuwe boeking</Btn>}
      </div>

      {saldo && (
        <Card className="p-3 flex flex-wrap gap-4 text-sm">
          <div><span style={{ color: C.inkSoft }}>Startsaldo {jaar}: </span><strong>{euro(saldo.startsaldo)}</strong></div>
          <div><span style={{ color: C.inkSoft }}>Mutaties: </span><strong style={{ color: saldo.mutaties >= 0 ? C.sageDeep : C.rose }}>{saldo.mutaties >= 0 ? '+' : ''}{euro(saldo.mutaties)}</strong></div>
          <div><span style={{ color: C.inkSoft }}>Eindsaldo: </span><strong>{euro(saldo.eindsaldo)}</strong></div>
        </Card>
      )}

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left border-b" style={{ borderColor: C.border }}>
              <th onClick={() => toggleSort('datum')} className="px-3 py-2 font-medium text-xs cursor-pointer select-none" style={{ color: C.inkSoft }}>Datum{sortIcon('datum')}</th>
              <th onClick={() => toggleSort('rekening')} className="px-3 py-2 font-medium text-xs cursor-pointer select-none" style={{ color: C.inkSoft }}>Rek.{sortIcon('rekening')}</th>
              <th onClick={() => toggleSort('grootboek')} className="px-3 py-2 font-medium text-xs cursor-pointer select-none" style={{ color: C.inkSoft }}>Grootboek{sortIcon('grootboek')}</th>
              <th onClick={() => toggleSort('omschrijving')} className="px-3 py-2 font-medium text-xs cursor-pointer select-none" style={{ color: C.inkSoft }}>Omschrijving{sortIcon('omschrijving')}</th>
              <th onClick={() => toggleSort('bedrag')} className="px-3 py-2 font-medium text-xs cursor-pointer select-none text-right" style={{ color: C.inkSoft }}>Bedrag{sortIcon('bedrag')}</th>
              <th></th>
            </tr>
            <tr className="border-b" style={{ borderColor: C.border }}>
              <td className="px-2 py-1"><input value={kolomFilter.datum} onChange={e => setKolomFilter(f => ({ ...f, datum: e.target.value }))} placeholder="filter" className="w-full text-xs rounded border px-1.5 py-0.5" style={inputStyle} /></td>
              <td></td>
              <td className="px-2 py-1"><input value={kolomFilter.grootboek} onChange={e => setKolomFilter(f => ({ ...f, grootboek: e.target.value }))} placeholder="filter" className="w-full text-xs rounded border px-1.5 py-0.5" style={inputStyle} /></td>
              <td></td>
              <td className="px-2 py-1"><input value={kolomFilter.bedrag} onChange={e => setKolomFilter(f => ({ ...f, bedrag: e.target.value }))} placeholder="filter" className="w-full text-xs rounded border px-1.5 py-0.5" style={inputStyle} /></td>
              <td></td>
            </tr>
          </thead>
          <tbody>
            {sorted.map(t => (
              <tr key={t.id} className="border-b last:border-0 hover:bg-black/[0.02]" style={{ borderColor: C.border }}>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>{fmtDate(t.datum)}</td>
                <td className="px-3 py-2" style={{ color: C.inkSoft }}>.{t.rekening}</td>
                <td className="px-3 py-2">{gbLabel(t)}</td>
                <td className="px-3 py-2">
                  {t.omschrijving}
                  {t.gekoppeldType && <Link2 size={12} className="inline ml-1.5 align-text-top" style={{ color: C.sageDeep }} />}
                </td>
                <td className="px-3 py-2 text-right font-medium" style={{ color: t.bedrag >= 0 ? C.sageDeep : C.rose }}>{t.bedrag >= 0 ? '+' : ''}{euro(t.bedrag)}</td>
                <td className="px-3 py-2">
                  {!readOnly && (
                    <div className="flex gap-1 justify-end">
                      <button onClick={() => { setEditing(t); setShowForm(true); }} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={14} /></button>
                      <button onClick={() => setDelId(t.id)} className="p-1.5 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={14} /></button>
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
          {sorted.length > 0 && (
            <tfoot>
              <tr className="border-t-2" style={{ borderColor: C.ink }}>
                <td colSpan={4} className="px-3 py-2 text-right font-semibold text-xs" style={{ color: C.inkSoft }}>Totaal ({sorted.length} boekingen)</td>
                <td className="px-3 py-2 text-right font-semibold" style={{ color: totaal >= 0 ? C.sageDeep : C.rose }}>{totaal >= 0 ? '+' : ''}{euro(totaal)}</td>
                <td></td>
              </tr>
            </tfoot>
          )}
        </table>
        {!sorted.length && <EmptyState icon={Wallet} text="Geen boekingen gevonden met deze filters." />}
      </Card>

      {showNewYear && (
        <PromptModal title="Nieuw boekjaar" label="Jaartal" placeholder="bv. 2028" onSave={v => addYear(Number(v))} onCancel={() => setShowNewYear(false)} />
      )}
      {showForm && <TxForm item={editing} accounts={accounts} members={members} contributies={contributies} workshops={workshops} inschrijvingen={inschrijvingen} onSave={save} onClose={() => { setShowForm(false); setEditing(null); }} />}
      {delId != null && (
        <ConfirmModal message={tx.find(t => t.id === delId)?.gekoppeldType ? 'Deze boeking verwijderen? Dit zet de gekoppelde contributie/inschrijving ook weer terug naar open.' : 'Deze boeking verwijderen?'} onConfirm={() => { remove(delId); setDelId(null); }} onCancel={() => setDelId(null)} />
      )}
    </div>
  );
}

function GrootboekPicker({ accounts, code, onPick }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState('');
  const gekozen = accounts.find(a => a.code === code);
  const term = q.trim().toLowerCase();
  const hits = accounts.filter(a => !term || a.code.includes(term) || a.naam.toLowerCase().includes(term)).slice(0, 30);
  return (
    <div className="relative">
      <button type="button" onClick={() => { setOpen(o => !o); setQ(''); }} className={`${inputCls} text-left flex items-center justify-between`} style={inputStyle}>
        <span>{gekozen ? `${gekozen.code} — ${gekozen.naam}` : '— kies grootboekrekening —'}</span>
        <ChevronDown size={14} style={{ color: C.inkSoft }} />
      </button>
      {open && (
        <div className="absolute left-0 right-0 mt-1 max-h-64 overflow-y-auto rounded-lg border shadow-lg z-50" style={{ background: C.card, borderColor: C.border }}>
          <input autoFocus value={q} onChange={e => setQ(e.target.value)} placeholder="Zoek op code of naam…" className="w-full px-3 py-2 text-sm border-b outline-none" style={{ borderColor: C.border }} />
          {hits.map(a => (
            <button key={a.code} type="button" onClick={() => { onPick(a.code); setOpen(false); }} className="w-full text-left px-3 py-1.5 text-sm hover:bg-black/5">
              {a.code} — {a.naam}
            </button>
          ))}
          {!hits.length && <p className="px-3 py-2 text-xs" style={{ color: C.inkSoft }}>Geen resultaten.</p>}
        </div>
      )}
    </div>
  );
}

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
  const huidigJaar = String(new Date().getFullYear());
  const jarenLijst = Array.from(new Set([...boekjaren.map(String), ...Object.keys(budget)])).sort();
  const [jaar, setJaar] = useState(jarenLijst.includes(huidigJaar) ? huidigJaar : (jarenLijst[jarenLijst.length - 1] || huidigJaar));
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

  function saveRow(data) {
    const nieuw = [...rows];
    if (data.idx != null) nieuw[data.idx] = { sectie: data.sectie, categorie: data.categorie, bedrag: Number(data.bedrag), notitie: data.notitie };
    else nieuw.push({ sectie: data.sectie, categorie: data.categorie, bedrag: Number(data.bedrag), notitie: data.notitie });
    setBudget({ ...budget, [jaar]: nieuw });
    onLog(`Begrotingsregel ${data.idx != null ? 'bewerkt' : 'toegevoegd'}: ${data.categorie} (${jaar})`, 'begroting');
    setShowForm(false);
  }
  function removeRow(idx) {
    const regel = rows[idx];
    setBudget({ ...budget, [jaar]: rows.filter((_, i) => i !== idx) });
    if (regel) onLog(`Begrotingsregel verwijderd: ${regel.categorie} (${jaar})`, 'begroting');
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
        {['Inkomsten', 'Uitgaven'].map(sectie => {
          const secRows = rows.filter(r => r.sectie === sectie);
          const begroot = secRows.reduce((s, r) => s + Number(r.bedrag || 0), 0);
          const werkelijk = secRows.reduce((s, r) => { const w = actualFor(r.categorie); return s + (w != null ? Math.abs(w) : 0); }, 0);
          return (
            <Card key={sectie} className="p-4">
              <p className="text-xs" style={{ color: C.inkSoft }}>{sectie} begroot</p>
              <p className="text-xl font-semibold" style={{ fontFamily: 'Fraunces, serif' }}>{euro(begroot)}</p>
              <p className="text-xs mt-1" style={{ color: C.inkSoft }}>Werkelijk (gekoppeld): {euro(werkelijk)}</p>
            </Card>
          );
        })}
        <Card className="p-4">
          <p className="text-xs" style={{ color: C.inkSoft }}>Saldo boekjaar {jaar}</p>
          {(() => { const s = saldoBoekjaar(tx, 'alle', Number(jaar)); return (
            <>
              <p className="text-xl font-semibold" style={{ fontFamily: 'Fraunces, serif' }}>{euro(s.eindsaldo)}</p>
              <p className="text-xs mt-1" style={{ color: C.inkSoft }}>Start: {euro(s.startsaldo)} · Mutatie: {s.mutaties >= 0 ? '+' : ''}{euro(s.mutaties)}</p>
            </>
          ); })()}
        </Card>
      </div>

      {['Inkomsten', 'Uitgaven'].map(sectie => {
        const secRows = rows.map((r, idx) => ({ ...r, idx })).filter(r => r.sectie === sectie);
        const totBegroot = secRows.reduce((s, r) => s + Number(r.bedrag || 0), 0);
        const totWerkelijk = secRows.reduce((s, r) => { const w = actualFor(r.categorie); return s + (w != null ? Math.abs(w) : 0); }, 0);
        const ongekoppeld = secRows.filter(r => !budgetGbCodes(r.categorie, begrotingKoppelingen).length);
        return (
          <div key={sectie}>
            <h3 className="font-semibold mb-2" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>{sectie}</h3>
            <Card className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left border-b" style={{ borderColor: C.border }}>
                    {['Categorie', 'Begroot', 'Werkelijk', 'Notitie', ''].map(h => <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {secRows.map(r => {
                    const w = actualFor(r.categorie);
                    return (
                      <tr key={r.idx} className="border-b last:border-0" style={{ borderColor: C.border }}>
                        <td className="px-3 py-2">{r.categorie}</td>
                        <td className="px-3 py-2">{euro(r.bedrag)}</td>
                        <td className="px-3 py-2" style={{ color: C.inkSoft }}>{w != null ? euro(Math.abs(w)) : '—'}</td>
                        <td className="px-3 py-2 text-xs" style={{ color: C.inkSoft }}>{r.notitie || ''}</td>
                        <td className="px-3 py-2">
                          {!readOnly && (
                            <div className="flex gap-1 justify-end">
                              <button onClick={() => { setShowForm({ ...r }); }} className="p-1 rounded hover:bg-black/5" style={{ color: C.inkSoft }}><Pencil size={13} /></button>
                              <button onClick={() => setDelIdx(r.idx)} className="p-1 rounded hover:bg-black/5" style={{ color: C.rose }}><Trash2 size={13} /></button>
                            </div>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr className="border-t-2" style={{ borderColor: C.ink }}>
                    <td className="px-3 py-2 font-semibold text-xs" style={{ color: C.inkSoft }}>Totaal</td>
                    <td className="px-3 py-2 font-semibold">{euro(totBegroot)}</td>
                    <td className="px-3 py-2 font-semibold" style={{ color: C.inkSoft }}>{euro(totWerkelijk)}</td>
                    <td colSpan={2} className="px-3 py-2 text-xs" style={{ color: C.inkSoft }}>
                      {ongekoppeld.length ? `Niet gekoppeld aan grootboek: ${ongekoppeld.map(r => r.categorie).join(', ')}` : ''}
                    </td>
                  </tr>
                </tfoot>
              </table>
              {!secRows.length && <EmptyState icon={PiggyBank} text={`Nog geen ${sectie.toLowerCase()} in de begroting van ${jaar}.`} />}
            </Card>
          </div>
        );
      })}

      {showNewYear && (
        <NewYearModal existing={jarenLijst.map(Number)} onSave={addYear} onClose={() => setShowNewYear(false)}
          label="boekjaar" copyFrom={Object.keys(budget).filter(y => budget[y] && budget[y].length)} />
      )}
      {showForm && (
        <BegrotingsregelForm item={typeof showForm === 'object' ? showForm : null} onSave={saveRow} onClose={() => setShowForm(false)} />
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

function NewYearModal({ existing, onSave, onClose, label, copyFrom }) {
  const volgende = existing.length ? Math.max(...existing) + 1 : new Date().getFullYear();
  const [jaar, setJaar] = useState(volgende);
  const [kopieerVan, setKopieerVan] = useState(copyFrom && copyFrom.length ? copyFrom[copyFrom.length - 1] : '');
  return (
    <Modal title={`Nieuw ${label}`} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Jaartal"><input type="number" className={inputCls} style={inputStyle} value={jaar} onChange={e => setJaar(Number(e.target.value))} /></Field>
        {copyFrom && copyFrom.length > 0 && (
          <Field label="Begroting kopiëren van (optioneel, bedragen worden op 0 gezet)">
            <select className={inputCls} style={inputStyle} value={kopieerVan} onChange={e => setKopieerVan(e.target.value)}>
              <option value="">— leeg beginnen —</option>
              {copyFrom.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
          </Field>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(jaar, kopieerVan)} disabled={existing.includes(jaar)}>Aanmaken</Btn>
      </div>
    </Modal>
  );
}

function BegrotingsregelForm({ item, onSave, onClose }) {
  const [f, setF] = useState(item ? { ...item } : { sectie: 'Inkomsten', categorie: '', bedrag: '', notitie: '' });
  function upd(k, v) { setF(s => ({ ...s, [k]: v })); }
  return (
    <Modal title={item ? 'Begrotingsregel bewerken' : 'Nieuwe begrotingsregel'} onClose={onClose}>
      <div className="space-y-3">
        <Field label="Sectie">
          <select className={inputCls} style={inputStyle} value={f.sectie} onChange={e => upd('sectie', e.target.value)}>
            <option value="Inkomsten">Inkomsten</option>
            <option value="Uitgaven">Uitgaven</option>
          </select>
        </Field>
        <Field label="Categorie"><input className={inputCls} style={inputStyle} value={f.categorie} onChange={e => upd('categorie', e.target.value)} /></Field>
        <Field label="Begroot bedrag"><input type="number" step="0.01" className={inputCls} style={inputStyle} value={f.bedrag} onChange={e => upd('bedrag', e.target.value)} /></Field>
        <Field label="Notitie (optioneel)"><input className={inputCls} style={inputStyle} value={f.notitie || ''} onChange={e => upd('notitie', e.target.value)} /></Field>
      </div>
      <div className="flex justify-end gap-2 mt-5">
        <Btn tone="ghost" onClick={onClose}>Annuleren</Btn>
        <Btn onClick={() => onSave(f)} disabled={!f.categorie}>Opslaan</Btn>
      </div>
    </Modal>
  );
}

/* =========================================================================
   RAPPORTAGE
========================================================================= */
function RapportageTab({ tx, accounts, members, workshops, inschrijvingen, budget, begrotingKoppelingen, logoHoogteCm }) {
  const years = Array.from(new Set(tx.map(t => t.jaar))).sort((a, b) => b - a);
  const huidigJaar = new Date().getFullYear();
  const [jaarNum, setJaarNum] = useState(years.includes(huidigJaar) ? huidigJaar : (years[0] || huidigJaar));
  const [maandFilter, setMaandFilter] = useState('alle');

  const jaarTx = tx.filter(t => t.jaar === jaarNum && (maandFilter === 'alle' || t.maand === Number(maandFilter)));
  const inkomsten = jaarTx.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0);
  const uitgaven = jaarTx.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0);

  const monthly = MONTH_NAMES.map((naam, i) => {
    const maand = i + 1;
    const rows = tx.filter(t => t.jaar === jaarNum && t.maand === maand);
    return { naam, Inkomsten: Math.round(rows.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0)), Uitgaven: Math.round(-rows.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0)) };
  });

  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const aoa = [['Financieel verslag', `${jaarNum}${maandFilter !== 'alle' ? ' - ' + MONTH_NAMES[Number(maandFilter) - 1] : ''}`], [],
      ['Datum', 'Rekening', 'Grootboek', 'Omschrijving', 'Bedrag'],
      ...jaarTx.map(t => [fmtDate(t.datum), '.' + t.rekening, gbLabel(t), t.omschrijving, t.bedrag]),
      [], ['Totaal inkomsten', '', '', '', inkomsten], ['Totaal uitgaven', '', '', '', uitgaven], ['Resultaat', '', '', '', inkomsten + uitgaven]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Financieel verslag');
    XLSX.writeFile(wb, `BladelsCreatief_Financieel_${jaarNum}.xlsx`);
  }
  function exportPdf() {
    downloadPrintableHtml({
      titel: `Financieel verslag ${jaarNum}`,
      subtitel: `Inkomsten: ${euro(inkomsten)} · Uitgaven: ${euro(uitgaven)} · Resultaat: ${euro(inkomsten + uitgaven)}`,
      header: ['Datum', 'Rekening', 'Grootboek', 'Omschrijving', 'Bedrag'],
      rows: jaarTx.map(t => [fmtDate(t.datum), '.' + t.rekening, gbLabel(t), t.omschrijving, euro(t.bedrag)]),
    });
  }
  function exportWord() {
    const gbLabelForTx = t => gbLabel(t);
    const map = {};
    jaarTx.forEach(t => { const k = gbLabelForTx(t) || 'Onbekend'; map[k] = (map[k] || 0) + t.bedrag; });
    const gbRows = Object.entries(map).sort((a, b) => b[1] - a[1]).map(([k, v]) => `<tr><td>${escapeHtml(k)}</td><td>${euro(v)}</td></tr>`).join('');
    const body = `${logoImgTag(logoHoogteCm)}<div class="bar"></div>
<h1>Financieel verslag ${jaarNum}</h1>
<p class="meta">Inkomsten: ${euro(inkomsten)} · Uitgaven: ${euro(uitgaven)} · Resultaat: ${euro(inkomsten + uitgaven)}</p>
<h2>Per grootboekrekening</h2>
<table><tr><th>Grootboekrekening</th><th>Bedrag</th></tr>${gbRows}</table>`;
    downloadWordDoc({ titel: `Financieel verslag ${jaarNum}`, filename: `BladelsCreatief_Financieel_${jaarNum}.doc`, bodyHtml: body });
  }

  function exportJaarverslagWord() {
    const jaarTxVoorVerslag = tx.filter(t => t.jaar === jaarNum);
    const inkomstenJ = jaarTxVoorVerslag.filter(t => t.bedrag > 0).reduce((s, t) => s + t.bedrag, 0);
    const uitgavenJ = jaarTxVoorVerslag.filter(t => t.bedrag < 0).reduce((s, t) => s + t.bedrag, 0);
    const actiefEindeJaar = members.filter(m => m.status === 'actief').length;
    const workshopsJaar = workshops.filter(w => (w.datums || []).some(d => d && d.startsWith(String(jaarNum))));
    const uniekeDeelnemers = new Set(inschrijvingen.filter(i => workshopsJaar.some(w => w.id === i.workshopId)).map(i => i.naam)).size;

    const bJaar = budget[String(jaarNum)] || [];
    const begrotingRows = bJaar.map(r => {
      const codes = budgetGbCodes(r.categorie, begrotingKoppelingen);
      const werkelijk = codes.length ? jaarTxVoorVerslag.filter(t => codes.includes(t.grootboek_code)).reduce((s, t) => s + t.bedrag, 0) : null;
      return `<tr><td>${escapeHtml(r.sectie)}</td><td>${escapeHtml(r.categorie)}</td><td>${euro(r.bedrag)}</td><td>${werkelijk == null ? '—' : euro(werkelijk)}</td></tr>`;
    }).join('');

    const body = `${logoImgTag(logoHoogteCm)}<div class="bar"></div>
<h1>Jaarverslag ${jaarNum}</h1>
<p class="meta">BladelsCreatief · gegenereerd op ${new Date().toLocaleDateString('nl-NL')}</p>
<h2>Leden</h2>
<p>${actiefEindeJaar} actieve leden aan het einde van de rapportageperiode.</p>
<h2>Workshops</h2>
<p>${workshopsJaar.length} workshop(s) in ${jaarNum}, met in totaal ${uniekeDeelnemers} unieke deelnemers.</p>
<h2>Financieel overzicht</h2>
<p>Inkomsten: ${euro(inkomstenJ)} · Uitgaven: ${euro(uitgavenJ)} · Resultaat: ${euro(inkomstenJ + uitgavenJ)}</p>
<h2>Begroting versus werkelijk</h2>
<table><tr><th>Sectie</th><th>Categorie</th><th>Begroot</th><th>Werkelijk</th></tr>${begrotingRows || '<tr><td colspan="4"><em>Geen begroting vastgelegd voor dit jaar.</em></td></tr>'}</table>`;
    downloadWordDoc({ titel: `Jaarverslag ${jaarNum}`, filename: `BladelsCreatief_Jaarverslag_${jaarNum}.doc`, bodyHtml: body });
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2">
          <select value={jaarNum} onChange={e => setJaarNum(Number(e.target.value))} className={inputCls} style={inputStyle}>
            {(years.length ? years : [huidigJaar]).map(j => <option key={j} value={j}>{j}</option>)}
          </select>
          <select value={maandFilter} onChange={e => setMaandFilter(e.target.value)} className={inputCls} style={inputStyle}>
            <option value="alle">Heel jaar</option>
            {MONTH_NAMES.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
          </select>
        </div>
        <div className="flex gap-2 flex-wrap">
          <Btn tone="outline" icon={Printer} onClick={exportPdf}>PDF</Btn>
          <Btn tone="outline" icon={FileText} onClick={exportWord}>Word</Btn>
          <Btn tone="sage" icon={Download} onClick={exportExcel}>Excel</Btn>
          <Btn tone="outline" icon={FileSpreadsheet} onClick={exportJaarverslagWord}>Compleet jaarverslag (Word)</Btn>
        </div>
      </div>

      <div className="grid sm:grid-cols-3 gap-3">
        <Card className="p-4"><p className="text-xs" style={{ color: C.inkSoft }}>Inkomsten</p><p className="text-xl font-semibold" style={{ fontFamily: 'Fraunces, serif', color: C.sageDeep }}>{euro(inkomsten)}</p></Card>
        <Card className="p-4"><p className="text-xs" style={{ color: C.inkSoft }}>Uitgaven</p><p className="text-xl font-semibold" style={{ fontFamily: 'Fraunces, serif', color: C.rose }}>{euro(uitgaven)}</p></Card>
        <Card className="p-4"><p className="text-xs" style={{ color: C.inkSoft }}>Resultaat</p><p className="text-xl font-semibold" style={{ fontFamily: 'Fraunces, serif' }}>{euro(inkomsten + uitgaven)}</p></Card>
      </div>

      <Card className="p-4">
        <h3 className="font-semibold mb-3" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Per maand · {jaarNum}</h3>
        <div style={{ height: 260 }}>
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={monthly}>
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
    </div>
  );
}

/* =========================================================================
   INSTELLINGEN
========================================================================= */
function InstellingenTab({ isVoorzitter, ingelogd, rolpermissies, setRolpermissies, beveiliging, setBeveiliging, standaarden, setStandaarden, watIsNieuw, setWatIsNieuw, logboek, prullenbak, setPrullenbak,
  members, setMembers, workshops, setWorkshops, inschrijvingen, setInschrijvingen, tx, setTx, boekjaren, setBoekjaren, accounts, setAccounts,
  begrotingKoppelingen, setBegrotingKoppelingen,
  dagdelen, setDagdelen,
  workshopSoorten, setWorkshopSoorten, agendapuntenVooraf, setAgendapuntenVooraf, agendapuntenAfsluitend, setAgendapuntenAfsluitend,
  budget, setBudget, actielijst, setActielijst, contributies, setContributies, vergaderingen, setVergaderingen,
  overigeActiviteiten, setOverigeActiviteiten,
  magLeden, magWorkshops, magFinancien, magVergaderingen,
  magLedenImporteren, magWorkshopsImporteren, magFinancienImporteren, backupData, onLog }) {
  const [sectie, setSectie] = useState('rollen');
  const [actieveAnderen, setActieveAnderen] = useState(null);

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
    if (entry.type === 'lid') setMembers(m => [...m, entry.data]);
    if (entry.type === 'workshop') setWorkshops(w => [...w, entry.data]);
    if (entry.type === 'inschrijving') setInschrijvingen(i => [...i, entry.data]);
    if (entry.type === 'boeking') setTx(t => [...t, entry.data]);
    if (entry.type === 'vergadering') setVergaderingen(v => [...v, entry.data]);
    if (entry.type === 'activiteit') setOverigeActiviteiten(a => [...a, entry.data]);
    setPrullenbak(prullenbak.filter(x => x.id !== entry.id));
    onLog(`Hersteld uit prullenbak: ${entry.type}`, 'instellingen');
  }
  function definitiefVerwijderen(id) {
    setPrullenbak(prullenbak.filter(x => x.id !== id));
  }

  async function checkActieveGebruikers() {
    setActieveAnderen(null);
    try {
      const res = await opslagLezen(STORAGE_KEYS.sessies, true);
      const sessies = res && res.value ? JSON.parse(res.value) : {};
      const grens = Date.now() - 3 * 60 * 1000;
      const eigenNaam = ingelogd ? fullName(ingelogd) : null;
      const anderen = Object.entries(sessies)
        .filter(([naam, tijd]) => naam !== eigenNaam && new Date(tijd).getTime() > grens)
        .map(([naam]) => naam);
      setActieveAnderen(anderen);
    } catch (e) {
      setActieveAnderen([]);
    }
  }
  useEffect(() => { if (sectie === 'backup') checkActieveGebruikers(); }, [sectie]);

  const [herstelData, setHerstelData] = useState(null);
  const [herstelFout, setHerstelFout] = useState('');
  const [herstelResultaat, setHerstelResultaat] = useState(false);
  const [toonBevestiging, setToonBevestiging] = useState(false);
  function handleHerstelFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setHerstelFout('');
    setHerstelResultaat(false);
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        setHerstelData(data);
      } catch (err) {
        setHerstelFout('Kon dit bestand niet lezen — is het een geldig BladelsCreatief-back-upbestand (.json)?');
      }
    };
    reader.readAsText(file);
  }
  function bevestigHerstel() {
    if (!herstelData) return;
    if (herstelData.members) setMembers(herstelData.members);
    if (herstelData.workshops) setWorkshops(herstelData.workshops);
    if (herstelData.inschrijvingen) setInschrijvingen(herstelData.inschrijvingen);
    if (herstelData.tx) setTx(herstelData.tx);
    if (herstelData.accounts) setAccounts(herstelData.accounts);
    if (herstelData.budget) setBudget(herstelData.budget);
    if (herstelData.boekjaren) setBoekjaren(herstelData.boekjaren);
    if (herstelData.vergaderingen) setVergaderingen(herstelData.vergaderingen);
    if (herstelData.actielijst) setActielijst(herstelData.actielijst);
    if (herstelData.contributies) setContributies(herstelData.contributies);
    if (herstelData.rolpermissies) setRolpermissies(herstelData.rolpermissies);
    if (herstelData.begrotingKoppelingen) setBegrotingKoppelingen(herstelData.begrotingKoppelingen);
    if (herstelData.overigeActiviteiten) setOverigeActiviteiten(herstelData.overigeActiviteiten);
    setHerstelResultaat(true);
    setToonBevestiging(false);
    onLog('Back-up hersteld', 'instellingen');
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

  return (
    <div className="flex flex-col sm:flex-row gap-4">
      <nav className="flex sm:flex-col gap-1 overflow-x-auto sm:w-56 flex-shrink-0">
        {SECTIES.map(s => {
          const Icon = s.icon;
          return (
            <button key={s.id} onClick={() => setSectie(s.id)}
              className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-medium text-left whitespace-nowrap"
              style={{ background: sectie === s.id ? C.paperDim : 'transparent', color: sectie === s.id ? C.clay : C.ink }}>
              <Icon size={15} /> {s.label}
            </button>
          );
        })}
      </nav>

      {sectie === 'rollen' && (
        <RolBeheer rolpermissies={rolpermissies} setRolpermissies={setRolpermissies} readOnly={!isVoorzitter} onLog={onLog} />
      )}

      {sectie === 'beveiliging' && (
        <BeveiligingBeheer beveiliging={beveiliging} setBeveiliging={setBeveiliging} readOnly={!isVoorzitter} onLog={onLog} />
      )}

      {sectie === 'standaarden' && (
        <StandaardenBeheer standaarden={standaarden} setStandaarden={setStandaarden} watIsNieuw={watIsNieuw} setWatIsNieuw={setWatIsNieuw} readOnly={!isVoorzitter} onLog={onLog} />
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
        <Card className="overflow-x-auto flex-1">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left border-b" style={{ borderColor: C.border }}>
                {['Tijdstip', 'Gebruiker', 'Gebied', 'Actie'].map(h => <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {logboek.map(l => (
                <tr key={l.id} className="border-b last:border-0" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2 text-xs" style={{ color: C.inkSoft }}>{new Date(l.tijdstip).toLocaleString('nl-NL')}</td>
                  <td className="px-3 py-2">{l.gebruiker}</td>
                  <td className="px-3 py-2"><Badge tone="muted">{l.gebied}</Badge></td>
                  <td className="px-3 py-2">{l.actie}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {!logboek.length && <EmptyState icon={History} text="Nog geen activiteit gelogd." />}
        </Card>
      )}

      {sectie === 'prullenbak' && (
        <Card className="overflow-x-auto flex-1">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left border-b" style={{ borderColor: C.border }}>
                {['Type', 'Omschrijving', 'Verwijderd op', 'Door', ''].map(h => <th key={h} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {prullenbak.map(entry => (
                <tr key={entry.id} className="border-b last:border-0" style={{ borderColor: C.border }}>
                  <td className="px-3 py-2"><Badge tone="muted">{entry.type}</Badge></td>
                  <td className="px-3 py-2">{entry.data.titel || entry.data.naam || fullName(entry.data) || entry.data.omschrijving || '—'}</td>
                  <td className="px-3 py-2 text-xs" style={{ color: C.inkSoft }}>{new Date(entry.verwijderdOp).toLocaleString('nl-NL')}</td>
                  <td className="px-3 py-2 text-xs" style={{ color: C.inkSoft }}>{entry.verwijderdDoor}</td>
                  <td className="px-3 py-2">
                    <div className="flex gap-2 justify-end">
                      <button onClick={() => herstel(entry)} className="text-xs underline flex items-center gap-1" style={{ color: C.sageDeep }}><RotateCcw size={12} /> herstel</button>
                      <button onClick={() => definitiefVerwijderen(entry.id)} className="text-xs underline" style={{ color: C.rose }}>definitief verwijderen</button>
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
        <div className="space-y-4 flex-1">
          <Card className="p-5">
            <h3 className="font-semibold mb-2" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Volledige back-up</h3>
            <p className="text-sm mb-3" style={{ color: C.inkSoft }}>
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
              <div className="mt-3">
                <Btn tone="dangerSolid" onClick={() => setToonBevestiging(true)}>Nu herstellen &amp; overschrijven</Btn>
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
        <ConfirmModal title="Weet je het zeker?" message="Dit overschrijft alle huidige gegevens in deze app-versie definitief met de inhoud van het geüploade back-upbestand."
          confirmLabel="Ja, overschrijven" onConfirm={bevestigHerstel} onCancel={() => setToonBevestiging(false)} />
      )}
    </div>
  );
}

function RolBeheer({ rolpermissies, setRolpermissies, readOnly, onLog }) {
  const [editing, setEditing] = useState(null);
  function toggleTab(regelId, tabId) {
    setRolpermissies(rolpermissies.map(r => {
      if (r.id !== regelId) return r;
      const has = r.tabs.includes(tabId);
      return { ...r, tabs: has ? r.tabs.filter(t => t !== tabId) : [...r.tabs, tabId] };
    }));
  }
  function updatePatroon(regelId, patroon) {
    setRolpermissies(rolpermissies.map(r => r.id === regelId ? { ...r, patroon } : r));
  }
  function addRegel() {
    setRolpermissies([...rolpermissies, { id: uid(rolpermissies), patroon: '', tabs: [] }]);
  }
  function removeRegel(id) {
    setRolpermissies(rolpermissies.filter(r => r.id !== id));
  }
  return (
    <Card className="p-4 flex-1 overflow-x-auto">
      <p className="text-xs mb-3" style={{ color: C.inkSoft }}>
        Elke regel herkent een functienaam (een deel van de tekst, niet hoofdlettergevoelig) en bepaalt welke tabbladen iemand met die functie mag bewerken. Dashboard, Rapportage en Instellingen zijn altijd voor iedereen toegankelijk.
      </p>
      <table className="w-full text-sm">
        <thead>
          <tr className="text-left border-b" style={{ borderColor: C.border }}>
            <th className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>Functie bevat</th>
            {EDITEERBARE_TABS.map(t => <th key={t.id} className="px-3 py-2 font-medium text-xs" style={{ color: C.inkSoft }}>{t.label}</th>)}
            <th></th>
          </tr>
        </thead>
        <tbody>
          {rolpermissies.map(r => (
            <tr key={r.id} className="border-b last:border-0" style={{ borderColor: C.border }}>
              <td className="px-3 py-2">
                <input disabled={readOnly} value={r.patroon} onChange={e => updatePatroon(r.id, e.target.value)} className="rounded border px-2 py-1 text-sm w-32" style={inputStyle} />
              </td>
              {EDITEERBARE_TABS.map(t => (
                <td key={t.id} className="px-3 py-2 text-center">
                  <input type="checkbox" disabled={readOnly} checked={r.tabs.includes(t.id)} onChange={() => toggleTab(r.id, t.id)} />
                </td>
              ))}
              <td className="px-3 py-2">{!readOnly && <button onClick={() => removeRegel(r.id)} style={{ color: C.rose }}><Trash2 size={14} /></button>}</td>
            </tr>
          ))}
        </tbody>
      </table>
      {!readOnly && <button onClick={addRegel} className="text-xs underline mt-3" style={{ color: C.clay }}>+ regel toevoegen</button>}
    </Card>
  );
}

function BeveiligingBeheer({ beveiliging, setBeveiliging, readOnly, onLog }) {
  function bijwerken(patch) {
    setBeveiliging({ ...beveiliging, ...patch });
    onLog(`Beveiligingsinstellingen gewijzigd: ${Object.entries(patch).map(([k, v]) => `${k}=${v}`).join(', ')}`, 'instellingen');
  }
  return (
    <div className="space-y-4 flex-1">
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

function StandaardenBeheer({ standaarden, setStandaarden, watIsNieuw, setWatIsNieuw, readOnly, onLog }) {
  const [nieuwText, setNieuwText] = useState(watIsNieuw.tekst);
  function bijwerken(patch) {
    setStandaarden({ ...standaarden, ...patch });
    onLog(`Standaarden gewijzigd: ${Object.entries(patch).map(([k, v]) => `${k}=${v}`).join(', ')}`, 'instellingen');
  }
  function toonAanIedereen() {
    setWatIsNieuw({ tekst: nieuwText, laatstBijgewerkt: new Date().toISOString() });
    onLog(`"Wat is nieuw"-melding bijgewerkt en opnieuw getoond aan iedereen`, 'instellingen');
  }
  return (
    <div className="space-y-4 flex-1">
      <Card className="p-5">
        <h3 className="font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Wat is nieuw</h3>
        <p className="text-sm mb-3" style={{ color: C.inkSoft }}>
          Deze tekst verschijnt als afsluitbare melding bovenaan het scherm bij iedereen die nog niet op "sluiten" heeft geklikt sinds de laatste keer dat je 'm hier hebt bijgewerkt en verstuurd. Handig om bijvoorbeeld een nieuwe functie kort toe te lichten. Laat leeg om niets te tonen.
        </p>
        <textarea disabled={readOnly} rows={3} className={inputCls} style={inputStyle} placeholder="bv. Nieuw: de Kalender toont nu ook de schildergroepen/dagdelen als je dat filter aanvinkt."
          value={nieuwText} onChange={e => setNieuwText(e.target.value)} />
        {!readOnly && (
          <div className="flex items-center gap-3 mt-2">
            <Btn onClick={toonAanIedereen} disabled={nieuwText === watIsNieuw.tekst}>Opslaan &amp; opnieuw tonen aan iedereen</Btn>
            {watIsNieuw.laatstBijgewerkt && <span className="text-xs" style={{ color: C.inkSoft }}>Laatst getoond sinds: {new Date(watIsNieuw.laatstBijgewerkt).toLocaleString('nl-NL')}</span>}
          </div>
        )}
      </Card>

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

      <Card className="p-5">
        <h3 className="font-semibold mb-1" style={{ fontFamily: 'Fraunces, serif', color: C.ink }}>Verenigingsgegevens</h3>
        <p className="text-sm mb-3" style={{ color: C.inkSoft }}>
          Gebruikt in de voettekst van de welkomstbrief en als standaard jaarcontributiebedrag bij het genereren van ledenbrieven (zie Leden). Hier aanpassen werkt overal meteen door, zonder dat er ergens een los document bijgewerkt hoeft te worden.
        </p>
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label="Website"><DebouncedField disabled={readOnly} value={standaarden.verenigingWebsite} onCommit={v => bijwerken({ verenigingWebsite: v })} /></Field>
          <Field label="E-mailadres"><DebouncedField disabled={readOnly} value={standaarden.verenigingEmail} onCommit={v => bijwerken({ verenigingEmail: v })} /></Field>
          <Field label="Adres"><DebouncedField disabled={readOnly} value={standaarden.verenigingAdres} onCommit={v => bijwerken({ verenigingAdres: v })} /></Field>
          <Field label="IBAN"><DebouncedField disabled={readOnly} value={standaarden.verenigingIban} onCommit={v => bijwerken({ verenigingIban: v })} /></Field>
          <Field label="Standaard jaarcontributie (€)"><input type="number" disabled={readOnly} className={inputCls} style={inputStyle} value={standaarden.jaarcontributie} onChange={e => bijwerken({ jaarcontributie: Number(e.target.value) })} /></Field>
        </div>
      </Card>
    </div>
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
    <div className="space-y-3 flex-1">
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
    <div className="space-y-3 flex-1">
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
    <div className="space-y-3 flex-1">
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
    <div className="space-y-3 flex-1">
      {readOnly && <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Je hebt geen bewerkrechten voor Workshops, en kunt workshop-soorten daarom alleen bekijken.</p>}
      <div className="flex items-center justify-between">
        <p className="text-xs" style={{ color: C.inkSoft }}>{items.length} {titel === 'workshop-soort' ? "workshop-soorten" : titel}</p>
        {!readOnly && <Btn icon={Plus} onClick={() => { setEditing(null); setShowForm(true); }}>{titel === 'workshop-soort' ? 'Workshop-soort toevoegen' : 'Toevoegen'}</Btn>}
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
                        <button onClick={() => { setEditing(item); setShowForm(true); }}
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
    <div className="space-y-4 flex-1">
      <p className="text-xs" style={{ color: C.inkSoft }}>
        Deze punten worden als aanvinkoptie voorgesteld bij het aanmaken van een nieuwe vergadering ("Standaard agendapunten toevoegen"). "Vooraf" komt bovenaan de agenda, "Afsluitend" onderaan.
      </p>
      {readOnly && <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Je hebt geen bewerkrechten voor Vergaderingen, en kunt dit daarom alleen bekijken.</p>}
      <Lijst label="Vooraf" items={vooraf} setItems={setVooraf} />
      <Lijst label="Afsluitend" items={afsluitend} setItems={setAfsluitend} />
    </div>
  );
}

function ImportSectie({ members, setMembers, workshops, setWorkshops, inschrijvingen, setInschrijvingen, tx, setTx, boekjaren, setBoekjaren, accounts, alleDagdelen, magLeden, magWorkshops, magFinancien, onLog }) {
  const [type, setType] = useState('leden');
  const [resultaat, setResultaat] = useState(null);
  const [fout, setFout] = useState('');

  function parseExcelDatum(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return d.toISOString().slice(0, 10);
    }
    const d = new Date(v);
    return isNaN(d) ? null : d.toISOString().slice(0, 10);
  }
  function rijNaarLid(row) {
    const dagdelen = alleDagdelen.filter(slot => row[slot] != null && row[slot] !== '');
    return {
      voornaam: row['Voornaam'] || row['voornaam'] || '', achternaam: row['Achternaam'] || row['achternaam'] || '',
      tussenvoegsel: row['Tussenvoegsel'] || '', email: row['E-mail'] || row['Email'] || row['email'] || '',
      telefoon: String(row['Telefoon'] || row['telefoon'] || ''), adres: row['Adres'] || '',
      postcode: row['Postcode'] || '', woonplaats: row['Woonplaats'] || '',
      gebdatum: parseExcelDatum(row['Geboortedatum']), lidsinds: parseExcelDatum(row['Lid sinds']),
      status: (row['Status'] || 'actief').toLowerCase(), functie: row['Functie'] || null, dagdelen,
    };
  }

  function importeerLeden(rows) {
    let toegevoegd = 0, bijgewerkt = 0;
    let nieuweLijst = [...members];
    rows.forEach(row => {
      const data = rijNaarLid(row);
      if (!data.voornaam && !data.achternaam) return;
      const naam = `${data.voornaam} ${data.tussenvoegsel} ${data.achternaam}`.replace(/\s+/g, ' ').trim().toLowerCase();
      const idx = nieuweLijst.findIndex(m => fullName(m).toLowerCase().replace(/\s+/g, ' ') === naam);
      if (idx >= 0) { nieuweLijst[idx] = { ...nieuweLijst[idx], ...data }; bijgewerkt++; }
      else { nieuweLijst.push({ ...data, id: uid(nieuweLijst) }); toegevoegd++; }
    });
    setMembers(nieuweLijst);
    onLog(`Leden geïmporteerd: ${toegevoegd} toegevoegd, ${bijgewerkt} bijgewerkt`, 'instellingen');
    setResultaat({ toegevoegd, bijgewerkt });
  }

  function importeerJeugd(rows) {
    let ws = workshops.find(w => w.titel.toLowerCase().includes('geïmporteerd jeugdatelier'));
    let nieuweWorkshops = [...workshops];
    if (!ws) {
      ws = { id: uid(nieuweWorkshops), titel: 'Geïmporteerd jeugdatelier', soort: 'Jeugdatelier', type: 'reeks', datums: [], dagdeel: '', locatie: '', bedrag: null, maxDeelnemers: null, status: 'open' };
      nieuweWorkshops.push(ws);
      setWorkshops(nieuweWorkshops);
    }
    let toegevoegd = 0, bijgewerkt = 0;
    let nieuweInschrijvingen = [...inschrijvingen];
    rows.forEach(row => {
      const naam = row['Naam'] || row['naam'] || '';
      if (!naam) return;
      const data = {
        workshopId: ws.id, herkomst: 'extern', naam, email: row['E-mail'] || '', telefoon: String(row['Telefoon'] || ''),
        leeftijd: row['Leeftijd'] ? Number(row['Leeftijd']) : null, groep: row['Groep'] || '',
        status: 'ingeschreven', betaald: false, bedrag: null, datumInschrijving: parseExcelDatum(row['Datum']),
      };
      const idx = nieuweInschrijvingen.findIndex(i => i.workshopId === ws.id && i.naam.toLowerCase() === naam.toLowerCase());
      if (idx >= 0) { nieuweInschrijvingen[idx] = { ...nieuweInschrijvingen[idx], ...data }; bijgewerkt++; }
      else { nieuweInschrijvingen.push({ ...data, id: uid(nieuweInschrijvingen) }); toegevoegd++; }
    });
    setInschrijvingen(nieuweInschrijvingen);
    onLog(`Jeugdatelier geïmporteerd: ${toegevoegd} toegevoegd, ${bijgewerkt} bijgewerkt`, 'instellingen');
    setResultaat({ toegevoegd, bijgewerkt });
  }

  function importeerFinancien(rows) {
    let toegevoegd = 0;
    const nieuwe = [];
    rows.forEach(row => {
      const datum = parseExcelDatum(row['Datum']);
      const bedrag = Number(row['Bedrag']);
      if (!datum || isNaN(bedrag)) return;
      const rekening = String(row['Rekening'] || '908').replace('.', '');
      const code = String(row['Grootboekcode'] || row['Grootboek'] || '');
      const acc = accounts.find(a => a.code === code);
      nieuwe.push({
        id: 0, rekening, jaar: new Date(datum).getFullYear(), maand: new Date(datum).getMonth() + 1, datum,
        grootboek_code: code, grootboek_naam: acc ? acc.naam : (row['Omschrijving grootboek'] || ''),
        bedrag, omschrijving: row['Omschrijving'] || '',
      });
    });
    let volgende = uid(tx);
    const metId = nieuwe.map(n => ({ ...n, id: volgende++ }));
    setTx([...tx, ...metId]);
    const jaren = Array.from(new Set(metId.map(t => t.jaar)));
    const nieuweJaren = jaren.filter(j => !boekjaren.includes(j));
    if (nieuweJaren.length) setBoekjaren([...boekjaren, ...nieuweJaren].sort((a, b) => a - b));
    toegevoegd = metId.length;
    onLog(`Financiën geïmporteerd: ${toegevoegd} boekingen toegevoegd`, 'instellingen');
    setResultaat({ toegevoegd, bijgewerkt: 0 });
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setFout('');
    setResultaat(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array', cellDates: false });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        if (type === 'leden') importeerLeden(rows);
        else if (type === 'jeugd') importeerJeugd(rows);
        else if (type === 'financien') importeerFinancien(rows);
      } catch (err) {
        setFout('Kon dit bestand niet verwerken — controleer of het een geldig Excel-bestand is met de juiste kolomnamen.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  }

  const magHuidige = { leden: magLeden, jeugd: magWorkshops, financien: magFinancien }[type];

  return (
    <div className="space-y-4 flex-1">
      <div className="flex gap-1">
        {[['leden', 'Leden'], ['jeugd', 'Jeugdatelier'], ['financien', 'Financiën']].map(([id, label]) => (
          <button key={id} onClick={() => { setType(id); setResultaat(null); setFout(''); }} className="px-3 py-1.5 rounded-lg text-sm font-medium border"
            style={{ background: type === id ? C.clay : 'transparent', color: type === id ? '#fff' : C.ink, borderColor: type === id ? C.clay : C.border }}>
            {label}
          </button>
        ))}
      </div>
      {!magHuidige && (
        <p className="text-xs rounded-lg px-3 py-2" style={{ background: C.paperDim, color: C.inkSoft }}>Je hebt geen bewerkrechten voor dit onderdeel, en kunt hier daarom niet importeren.</p>
      )}
      {magHuidige && (
        <>
          <p className="text-sm" style={{ color: C.inkSoft }}>
            Upload een Excel-bestand (.xlsx). Bestaande gegevens worden samengevoegd op naam — niets gaat verloren, dubbele namen worden bijgewerkt in plaats van dubbel toegevoegd.
          </p>
          <input type="file" accept=".xlsx,.xls" onChange={handleFile} className="text-sm" />
          {fout && <p className="text-xs mt-2" style={{ color: C.rose }}>{fout}</p>}
          {resultaat && (
            <div className="mt-2 rounded-lg px-3 py-2 text-sm flex items-center gap-2" style={{ background: C.paperDim, color: C.sageDeep }}>
              <CheckCircle2 size={15} /> {resultaat.toegevoegd} toegevoegd, {resultaat.bijgewerkt} bijgewerkt.
            </div>
          )}
        </>
      )}
    </div>
  );
}
