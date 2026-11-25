# -*- coding: utf-8 -*-
# Trivia Challenge batch #7. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which purple vegetable is often grilled or used in ratatouille?', 'de': 'Welches violette Gemuese wird oft gegrillt oder in Ratatouille verwendet?', 'es': '¿Que verdura morada se asa a menudo o se usa en ratatouille?', 'nl': 'Welke paarse groente wordt vaak gegrild of gebruikt in ratatouille?', 'sv': 'Vilken lila gronsak grillas ofta eller anvands i ratatouille?', 'fr': "Quel legume violet est souvent grille ou utilise dans la ratatouille ?", 'da': 'Hvilken lilla grøntsag grilles ofte eller bruges i ratatouille?'},
 'opt1': {'en': 'Eggplant', 'de': 'Aubergine', 'es': 'Berenjena', 'nl': 'Aubergine', 'sv': 'Aubergine', 'fr': "L'aubergine", 'da': 'Aubergine'},
 'opt2': {'en': 'Cabbage', 'de': 'Kohl', 'es': 'Repollo', 'nl': 'Kool', 'sv': 'Kal', 'fr': 'Le chou', 'da': 'Kål'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the bullet train (Shinkansen) network?', 'de': 'In welchem Land befindet sich das Shinkansen-Hochgeschwindigkeitszugnetz?', 'es': '¿En que pais se encuentra la red del tren bala (Shinkansen)?', 'nl': 'In welk land bevindt zich het Shinkansen-hogesnelheidstreinnetwerk?', 'sv': 'I vilket land finns Shinkansen-snabbtagsnatet?', 'fr': "Dans quel pays se trouve le reseau de trains a grande vitesse (Shinkansen) ?", 'da': 'Hvilket land huser Shinkansen-hurtigtognettet?'},
 'opt1': {'en': 'South Korea', 'de': 'Suedkorea', 'es': 'Corea del Sur', 'nl': 'Zuid-Korea', 'sv': 'Sydkorea', 'fr': 'la Coree du Sud', 'da': 'Sydkorea'},
 'opt2': {'en': 'Japan', 'de': 'Japan', 'es': 'Japon', 'nl': 'Japan', 'sv': 'Japan', 'fr': 'le Japon', 'da': 'Japan'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic film features a boy raised by wolves in the jungle?', 'de': 'In welchem Filmklassiker wird ein Junge von Woelfen im Dschungel aufgezogen?', 'es': '¿Que pelicula clasica presenta a un niño criado por lobos en la selva?', 'nl': 'In welke klassieke film wordt een jongen opgevoed door wolven in de jungle?', 'sv': 'I vilken klassisk film uppfostras en pojke av vargar i djungeln?', 'fr': "Quel film classique met en scene un garçon eleve par des loups dans la jungle ?", 'da': 'Hvilken klassisk film har en dreng, der er opdraget af ulve i junglen?'},
 'opt1': {'en': 'Tarzan', 'de': 'Tarzan', 'es': 'Tarzan', 'nl': 'Tarzan', 'sv': 'Tarzan', 'fr': 'Tarzan', 'da': 'Tarzan'},
 'opt2': {'en': 'The Jungle Book', 'de': 'Das Dschungelbuch', 'es': 'El libro de la selva', 'nl': 'Het Jungle Boek', 'sv': 'Djungelboken', 'fr': 'Le Livre de la jungle', 'da': 'Junglebogen'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which pet rodent runs on a wheel and is often kept in a cage?', 'de': 'Welches Haustier-Nagetier laeuft auf einem Rad und wird oft in einem Kaefig gehalten?', 'es': '¿Que roedor mascota corre en una rueda y a menudo se mantiene en una jaula?', 'nl': 'Welk knaagdier huisdier rent op een wiel en wordt vaak in een kooi gehouden?', 'sv': 'Vilken sallskapsgnagare springer pa ett hjul och halls ofta i en bur?', 'fr': "Quel rongeur de compagnie court sur une roue et est souvent garde en cage ?", 'da': 'Hvilket kæledyrsgnaver løber på et hjul og holdes ofte i et bur?'},
 'opt1': {'en': 'Ferret', 'de': 'Frettchen', 'es': 'Huron', 'nl': 'Fret', 'sv': 'Iller', 'fr': 'Furet', 'da': 'Fritte'},
 'opt2': {'en': 'Hamster', 'de': 'Hamster', 'es': 'Hamster', 'nl': 'Hamster', 'sv': 'Hamster', 'fr': 'Hamster', 'da': 'Hamster'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which orange soup is traditionally made from squash or pumpkin?', 'de': 'Welche orangefarbene Suppe wird traditionell aus Kuerbis hergestellt?', 'es': '¿Que sopa naranja se hace tradicionalmente con calabaza?', 'nl': 'Welke oranje soep wordt traditioneel gemaakt van pompoen?', 'sv': 'Vilken orange soppa gors traditionellt av pumpa?', 'fr': "Quelle soupe orange est traditionnellement faite avec de la courge ou de la citrouille ?", 'da': 'Hvilken orange suppe er traditionelt lavet af squash eller græskar?'},
 'opt1': {'en': 'Butternut squash soup', 'de': 'Butternusskuerbis-Suppe', 'es': 'Sopa de calabaza butternut', 'nl': 'Butternutpompoensoep', 'sv': 'Butternutpumpasoppa', 'fr': 'Soupe de courge butternut', 'da': 'Butternut squash-suppe'},
 'opt2': {'en': 'Tomato soup', 'de': 'Tomatensuppe', 'es': 'Sopa de tomate', 'nl': 'Tomatensoep', 'sv': 'Tomatsoppa', 'fr': 'Soupe de tomate', 'da': 'Tomatsuppe'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the Alps mountain range (shared with others)?', 'de': 'In welchem Land liegt (unter anderem) das Alpengebirge?', 'es': '¿En que pais esta (compartida con otros) la cordillera de los Alpes?', 'nl': 'In welk land ligt (gedeeld met andere landen) het Alpengebergte?', 'sv': 'I vilket land ligger (delat med andra) Alpernas bergskedja?', 'fr': "Dans quel pays se trouve (partagee avec d'autres) la chaine des Alpes ?", 'da': 'Hvilket land huser (delt med andre) Alperne?'},
 'opt1': {'en': 'Switzerland', 'de': 'Schweiz', 'es': 'Suiza', 'nl': 'Zwitserland', 'sv': 'Schweiz', 'fr': 'la Suisse', 'da': 'Schweiz'},
 'opt2': {'en': 'Netherlands', 'de': 'Niederlande', 'es': 'Paises Bajos', 'nl': 'Nederland', 'sv': 'Nederlanderna', 'fr': 'les Pays-Bas', 'da': 'Nederlandene'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which U.S. president was assassinated in Ford's Theatre?", 'de': 'Welcher US-Praesident wurde im Ford Theatre ermordet?', 'es': '¿Que presidente de EE.UU. fue asesinado en el Teatro Ford?', 'nl': 'Welke Amerikaanse president werd vermoord in het Ford Theatre?', 'sv': 'Vilken amerikansk president mordades i Fords teater?', 'fr': "Quel president americain a ete assassine au Ford's Theatre ?", 'da': 'Hvilken amerikansk præsident blev myrdet i Ford’s Theatre?'},
 'opt1': {'en': 'James Garfield', 'de': 'James Garfield', 'es': 'James Garfield', 'nl': 'James Garfield', 'sv': 'James Garfield', 'fr': 'James Garfield', 'da': 'James Garfield'},
 'opt2': {'en': 'Abraham Lincoln', 'de': 'Abraham Lincoln', 'es': 'Abraham Lincoln', 'nl': 'Abraham Lincoln', 'sv': 'Abraham Lincoln', 'fr': 'Abraham Lincoln', 'da': 'Abraham Lincoln'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves knitting yarn into fabric using two needles?', 'de': 'Bei welchem Handwerk wird Garn mit zwei Nadeln zu Stoff gestrickt?', 'es': '¿Que manualidad implica tejer hilo en tela usando dos agujas?', 'nl': 'Bij welk handwerk brei je garen tot stof met twee naalden?', 'sv': 'Vilket hantverk innebar att sticka garn till tyg med tva stickor?', 'fr': "Quel art manuel consiste a tricoter du fil en tissu avec deux aiguilles ?", 'da': 'Hvilket håndværk involverer at strikke garn til stof med to pinde?'},
 'opt1': {'en': 'Knitting', 'de': 'Stricken', 'es': 'Tejido de punto', 'nl': 'Breien', 'sv': 'Stickning', 'fr': 'Le tricot', 'da': 'Strikning'},
 'opt2': {'en': 'Crochet', 'de': 'Haekeln', 'es': 'Ganchillo', 'nl': 'Haken', 'sv': 'Virkning', 'fr': 'Le crochet', 'da': 'Hækling'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which woodwind instrument is played by blowing across a hole near one end?', 'de': 'Welches Holzblasinstrument wird gespielt, indem man ueber ein Loch nahe einem Ende blaest?', 'es': '¿Que instrumento de viento madera se toca soplando a traves de un agujero cerca de un extremo?', 'nl': 'Welk houtblazersinstrument bespeel je door over een gat nabij een uiteinde te blazen?', 'sv': 'Vilket trablasinstrument spelas genom att blasa over ett hal nara ena anden?', 'fr': "Quel instrument a vent en bois se joue en soufflant sur un trou pres d'une extremite ?", 'da': 'Hvilket træblæserinstrument spilles ved at blæse over et hul nær den ene ende?'},
 'opt1': {'en': 'Flute', 'de': 'Floete', 'es': 'Flauta', 'nl': 'Fluit', 'sv': 'Flojt', 'fr': 'Flute', 'da': 'Fløjte'},
 'opt2': {'en': 'Clarinet', 'de': 'Klarinette', 'es': 'Clarinete', 'nl': 'Klarinet', 'sv': 'Klarinett', 'fr': 'Clarinette', 'da': 'Klarinet'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which season is typically the hottest of the year?', 'de': 'Welche Jahreszeit ist typischerweise die heisseste des Jahres?', 'es': '¿Que estacion es tipicamente la mas calurosa del año?', 'nl': 'Welk seizoen is doorgaans het warmste van het jaar?', 'sv': 'Vilken arstid ar vanligtvis den varmaste pa aret?', 'fr': "Quelle saison est typiquement la plus chaude de l'annee ?", 'da': 'Hvilken årstid er typisk den varmeste på året?'},
 'opt1': {'en': 'Spring', 'de': 'Fruehling', 'es': 'Primavera', 'nl': 'Lente', 'sv': 'Var', 'fr': 'Printemps', 'da': 'Forår'},
 'opt2': {'en': 'Summer', 'de': 'Sommer', 'es': 'Verano', 'nl': 'Zomer', 'sv': 'Sommar', 'fr': 'Ete', 'da': 'Sommer'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which classic film features toys that come to life when humans aren't looking?", 'de': 'In welchem Filmklassiker werden Spielzeuge lebendig, wenn Menschen nicht hinsehen?', 'es': '¿Que pelicula clasica presenta juguetes que cobran vida cuando los humanos no miran?', 'nl': 'In welke klassieke film komen speelgoed tot leven wanneer mensen niet kijken?', 'sv': 'I vilken klassisk film blir leksaker levande nar manniskor inte tittar?', 'fr': "Quel film classique met en scene des jouets qui prennent vie quand les humains ne regardent pas ?", 'da': 'Hvilken klassisk film har legetøj, der bliver levende, når mennesker ikke ser på?'},
 'opt1': {'en': 'Toy Story', 'de': 'Toy Story', 'es': 'Toy Story', 'nl': 'Toy Story', 'sv': 'Toy Story', 'fr': 'Toy Story', 'da': 'Toy Story'},
 'opt2': {'en': 'The Indian in the Cupboard', 'de': 'Der Indianer im Schrank', 'es': 'El indio de la alacena', 'nl': 'De Indiaan in de Kast', 'sv': 'Indianen i skapet', 'fr': "L'Indien du placard", 'da': 'Indianeren i skabet'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which U.S. state is known as the 'Golden State'?", 'de': "Welcher US-Bundesstaat ist als 'Golden State' bekannt?", 'es': "¿Que estado de EE.UU. es conocido como el 'Estado Dorado'?", 'nl': "Welke Amerikaanse staat staat bekend als de 'Golden State'?", 'sv': "Vilken amerikansk delstat ar kand som 'Golden State'?", 'fr': "Quel État americain est connu comme le 'Golden State' (État dore) ?", 'da': "Hvilken amerikansk delstat er kendt som 'Golden State'?"},
 'opt1': {'en': 'Nevada', 'de': 'Nevada', 'es': 'Nevada', 'nl': 'Nevada', 'sv': 'Nevada', 'fr': 'le Nevada', 'da': 'Nevada'},
 'opt2': {'en': 'California', 'de': 'Kalifornien', 'es': 'California', 'nl': 'Californie', 'sv': 'Kalifornien', 'fr': 'la Californie', 'da': 'Californien'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which red fruit is used to make traditional strawberry shortcake?', 'de': 'Welche rote Frucht wird fuer traditionellen Erdbeerkuchen verwendet?', 'es': '¿Que fruta roja se usa para hacer el tradicional pastel de fresa?', 'nl': 'Welk rood fruit wordt gebruikt om traditionele aardbeientaart te maken?', 'sv': 'Vilken rod frukt anvands for att gora traditionell jordgubbstarta?', 'fr': "Quel fruit rouge est utilise pour faire le traditionnel shortcake aux fraises ?", 'da': 'Hvilken rød frugt bruges til at lave traditionel jordbærkage?'},
 'opt1': {'en': 'Strawberry', 'de': 'Erdbeere', 'es': 'Fresa', 'nl': 'Aardbei', 'sv': 'Jordgubbe', 'fr': 'La fraise', 'da': 'Jordbær'},
 'opt2': {'en': 'Raspberry', 'de': 'Himbeere', 'es': 'Frambuesa', 'nl': 'Framboos', 'sv': 'Hallon', 'fr': 'La framboise', 'da': 'Hindbær'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which sport involves swimming, cycling, and running in sequence?', 'de': 'Bei welcher Sportart schwimmt, radelt und laeuft man nacheinander?', 'es': '¿Que deporte implica nadar, andar en bicicleta y correr en secuencia?', 'nl': 'Bij welke sport zwem, fiets en ren je achter elkaar?', 'sv': 'Vilken sport innebar att simma, cykla och springa i foljd?', 'fr': "Quel sport implique de nager, faire du velo et courir a la suite ?", 'da': 'Hvilken sport involverer at svømme, cykle og løbe i træk?'},
 'opt1': {'en': 'Duathlon', 'de': 'Duathlon', 'es': 'Duatlon', 'nl': 'Duatlon', 'sv': 'Duathlon', 'fr': 'Le duathlon', 'da': 'Duatlon'},
 'opt2': {'en': 'Triathlon', 'de': 'Triathlon', 'es': 'Triatlon', 'nl': 'Triatlon', 'sv': 'Triathlon', 'fr': 'Le triathlon', 'da': 'Triatlon'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which large aquatic reptile is native to the southeastern United States and has a rounded, U-shaped snout?', 'de': 'Welches grosse Wasserreptil ist im Suedosten der USA heimisch und hat eine abgerundete, U-foermige Schnauze?', 'es': '¿Que gran reptil acuatico es nativo del sureste de Estados Unidos y tiene un hocico redondeado en forma de U?', 'nl': 'Welk groot waterreptiel is inheems in het zuidoosten van de VS en heeft een ronde, U-vormige snuit?', 'sv': 'Vilket stort vattenlevande reptil ar inhemskt i sydostra USA och har en rundad, U-formad noshugg?', 'fr': "Quel grand reptile aquatique est originaire du sud-est des États-Unis et a un museau arrondi en forme de U ?", 'da': 'Hvilket stort vandkrybdyr er hjemmehørende i det sydøstlige USA og har en rund, U-formet snude?'},
 'opt1': {'en': 'Alligator', 'de': 'Alligator', 'es': 'Caiman/Aligator', 'nl': 'Alligator', 'sv': 'Alligator', 'fr': 'Alligator', 'da': 'Alligator'},
 'opt2': {'en': 'Crocodile', 'de': 'Krokodil', 'es': 'Cocodrilo', 'nl': 'Krokodil', 'sv': 'Krokodil', 'fr': 'Crocodile', 'da': 'Krokodille'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which holiday involves wearing costumes and parading with floats in New Orleans?', 'de': 'Bei welchem Feiertag in New Orleans traegt man Kostueme und paradiert mit Festwagen?', 'es': '¿Que festividad implica llevar disfraces y desfilar con carrozas en Nueva Orleans?', 'nl': 'Bij welke feestdag in New Orleans draag je kostuums en paradeer je met praalwagens?', 'sv': 'Vilken hogtid i New Orleans innebar att bara kostymer och parad med festvagnar?', 'fr': "Quelle fete de La Nouvelle-Orleans implique de porter des costumes et de defiler avec des chars ?", 'da': 'Hvilken helligdag i New Orleans involverer at bære kostumer og parade med festvogne?'},
 'opt1': {'en': 'Mardi Gras', 'de': 'Mardi Gras', 'es': 'Mardi Gras', 'nl': 'Mardi Gras', 'sv': 'Mardi Gras', 'fr': 'Mardi gras', 'da': 'Fastelavn (Mardi Gras)'},
 'opt2': {'en': 'Halloween', 'de': 'Halloween', 'es': 'Halloween', 'nl': 'Halloween', 'sv': 'Halloween', 'fr': 'Halloween', 'da': 'Halloween'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which continent is home to the Ganges and Mekong rivers?', 'de': 'Auf welchem Kontinent liegen der Ganges und der Mekong?', 'es': '¿En que continente estan los rios Ganges y Mekong?', 'nl': 'Op welk continent liggen de rivieren de Ganges en de Mekong?', 'sv': 'Pa vilken kontinent ligger floderna Ganges och Mekong?', 'fr': "Sur quel continent se trouvent le Gange et le Mekong ?", 'da': 'Hvilket kontinent huser floderne Ganges og Mekong?'},
 'opt1': {'en': 'Africa', 'de': 'Afrika', 'es': 'Africa', 'nl': 'Afrika', 'sv': 'Afrika', 'fr': "l'Afrique", 'da': 'Afrika'},
 'opt2': {'en': 'Asia', 'de': 'Asien', 'es': 'Asia', 'nl': 'Azie', 'sv': 'Asien', 'fr': "l'Asie", 'da': 'Asien'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which author wrote 'Where the Wild Things Are'?", 'de': "Welcher Autor schrieb 'Wo die wilden Kerle wohnen'?", 'es': "¿Que autor escribio 'Donde viven los monstruos'?", 'nl': "Welke auteur schreef 'Waar de Wilde Dingen Wonen'?", 'sv': "Vilken forfattare skrev 'Dar vildarna bor'?", 'fr': "Quel auteur a ecrit 'Max et les maximonstres' ?", 'da': "Hvilken forfatter skrev 'Hvor de vilde ting bor'?"},
 'opt1': {'en': 'Shel Silverstein', 'de': 'Shel Silverstein', 'es': 'Shel Silverstein', 'nl': 'Shel Silverstein', 'sv': 'Shel Silverstein', 'fr': 'Shel Silverstein', 'da': 'Shel Silverstein'},
 'opt2': {'en': 'Maurice Sendak', 'de': 'Maurice Sendak', 'es': 'Maurice Sendak', 'nl': 'Maurice Sendak', 'sv': 'Maurice Sendak', 'fr': 'Maurice Sendak', 'da': 'Maurice Sendak'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic toy is a set of colorful plastic blocks that snap together?', 'de': 'Welches klassische Spielzeug ist eine Reihe bunter Kunststoffbloecke, die zusammenschnappen?', 'es': '¿Que juguete clasico es un conjunto de bloques de plastico coloridos que encajan entre si?', 'nl': 'Welk klassiek speelgoed is een set kleurrijke plastic blokjes die in elkaar klikken?', 'sv': 'Vilken klassisk leksak ar en uppsattning fargglada plastklossar som klickar ihop?', 'fr': "Quel jouet classique est un ensemble de blocs en plastique colores qui s'emboitent ?", 'da': 'Hvilket klassisk legetøj er et sæt farverige plastikklodser, der klikker sammen?'},
 'opt1': {'en': 'Lincoln Logs', 'de': 'Lincoln Logs', 'es': 'Lincoln Logs', 'nl': 'Lincoln Logs', 'sv': 'Lincoln Logs', 'fr': 'Lincoln Logs', 'da': 'Lincoln Logs'},
 'opt2': {'en': 'LEGO', 'de': 'LEGO', 'es': 'LEGO', 'nl': 'LEGO', 'sv': 'LEGO', 'fr': 'LEGO', 'da': 'LEGO'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which black spice is ground and commonly found on dining tables?', 'de': 'Welches schwarze Gewuerz wird gemahlen und ist oft auf Esstischen zu finden?', 'es': '¿Que especia negra se muele y se encuentra comunmente en las mesas de comedor?', 'nl': 'Welke zwarte specerij wordt gemalen en vaak op eettafels gevonden?', 'sv': 'Vilken svart krydda mals och finns ofta pa matbord?', 'fr': "Quelle epice noire est moulue et se trouve couramment sur les tables a manger ?", 'da': 'Hvilket sort krydderi males og findes ofte på spiseborde?'},
 'opt1': {'en': 'Black pepper', 'de': 'Schwarzer Pfeffer', 'es': 'Pimienta negra', 'nl': 'Zwarte peper', 'sv': 'Svartpeppar', 'fr': 'Poivre noir', 'da': 'Sort peber'},
 'opt2': {'en': 'Cumin', 'de': 'Kreuzkuemmel', 'es': 'Comino', 'nl': 'Komijn', 'sv': 'Spiskummin', 'fr': 'Cumin', 'da': 'Spidskommen'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. holiday celebrates workers and falls on the first Monday of September?', 'de': 'Welcher US-Feiertag ehrt Arbeitnehmer und faellt auf den ersten Montag im September?', 'es': '¿Que festividad de EE.UU. celebra a los trabajadores y cae el primer lunes de septiembre?', 'nl': 'Welke Amerikaanse feestdag viert werknemers en valt op de eerste maandag van september?', 'sv': 'Vilken amerikansk hogtid firar arbetare och infaller den forsta mandagen i september?', 'fr': "Quelle fete americaine celebre les travailleurs et tombe le premier lundi de septembre ?", 'da': 'Hvilken amerikansk helligdag fejrer arbejdere og falder på den første mandag i september?'},
 'opt1': {'en': 'Memorial Day', 'de': 'Memorial Day', 'es': 'Dia de los Caidos', 'nl': 'Memorial Day', 'sv': 'Memorial Day', 'fr': 'Le Memorial Day', 'da': 'Memorial Day'},
 'opt2': {'en': 'Labor Day', 'de': 'Labor Day', 'es': 'Dia del Trabajo', 'nl': 'Labor Day', 'sv': 'Labor Day', 'fr': 'La fete du Travail', 'da': 'Labor Day'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which decade is associated with the rise of boy bands like the Backstreet Boys?', 'de': 'Mit welchem Jahrzehnt wird der Aufstieg von Boybands wie den Backstreet Boys assoziiert?', 'es': '¿Con que decada se asocia el auge de las boy bands como los Backstreet Boys?', 'nl': 'Met welk decennium wordt de opkomst van boybands zoals de Backstreet Boys geassocieerd?', 'sv': 'Vilket artionde forknippas med framvaxten av pojkband som Backstreet Boys?', 'fr': "A quelle decennie associe-t-on la montee des boys bands comme les Backstreet Boys ?", 'da': 'Hvilket årti forbindes med fremkomsten af drengebands som Backstreet Boys?'},
 'opt1': {'en': '1990s', 'de': '1990er', 'es': 'los años 90', 'nl': 'jaren 90', 'sv': '1990-talet', 'fr': 'les annees 1990', 'da': '1990’erne'},
 'opt2': {'en': '1980s', 'de': '1980er', 'es': 'los años 80', 'nl': 'jaren 80', 'sv': '1980-talet', 'fr': 'les annees 1980', 'da': '1980’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient Colosseum?', 'de': 'In welchem Land liegt das antike Kolosseum?', 'es': '¿En que pais esta el antiguo Coliseo?', 'nl': 'In welk land ligt het oude Colosseum?', 'sv': 'I vilket land ligger det antika Colosseum?', 'fr': "Dans quel pays se trouve l'ancien Colisee ?", 'da': 'Hvilket land huser det antikke Colosseum?'},
 'opt1': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'},
 'opt2': {'en': 'Italy', 'de': 'Italien', 'es': 'Italia', 'nl': 'Italie', 'sv': 'Italien', 'fr': "l'Italie", 'da': 'Italien'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which farm animal is known for neighing and being ridden?', 'de': 'Welches Nutztier wiehert und wird geritten?', 'es': '¿Que animal de granja relincha y se monta?', 'nl': 'Welk boerderijdier hinnikt en wordt bereden?', 'sv': 'Vilket lantbruksdjur gnaggar och rids pa?', 'fr': "Quel animal de ferme hennit et se monte ?", 'da': 'Hvilket landbrugsdyr vrinsker og rides på?'},
 'opt1': {'en': 'Horse', 'de': 'Pferd', 'es': 'Caballo', 'nl': 'Paard', 'sv': 'Hast', 'fr': 'Cheval', 'da': 'Hest'},
 'opt2': {'en': 'Donkey', 'de': 'Esel', 'es': 'Burro', 'nl': 'Ezel', 'sv': 'Asna', 'fr': 'Ane', 'da': 'Æsel'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which gemstone is the birthstone for those born in June?', 'de': 'Welcher Edelstein ist der Geburtsstein fuer im Juni Geborene?', 'es': '¿Que piedra preciosa es la piedra de nacimiento para los nacidos en junio?', 'nl': 'Welke edelsteen is de geboortesteen voor mensen geboren in juni?', 'sv': 'Vilken adelsten ar fodelsestenen for de fodda i juni?', 'fr': "Quelle pierre precieuse est la pierre de naissance de juin ?", 'da': 'Hvilken ædelsten er fødselsstenen for dem, der er født i juni?'},
 'opt1': {'en': 'Pearl', 'de': 'Perle', 'es': 'Perla', 'nl': 'Parel', 'sv': 'Parla', 'fr': 'Perle', 'da': 'Perle'},
 'opt2': {'en': 'Opal', 'de': 'Opal', 'es': 'Opalo', 'nl': 'Opaal', 'sv': 'Opal', 'fr': 'Opale', 'da': 'Opal'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic game involves stacking wooden blocks until the tower falls?', 'de': 'Bei welchem klassischen Spiel stapelt man Holzbloecke, bis der Turm faellt?', 'es': '¿Que juego clasico consiste en apilar bloques de madera hasta que la torre cae?', 'nl': 'Bij welk klassiek spel stapel je houten blokjes tot de toren valt?', 'sv': 'Vilket klassiskt spel innebar att stapla traklossar tills tornet faller?', 'fr': "Quel jeu classique consiste a empiler des blocs de bois jusqu'a ce que la tour tombe ?", 'da': 'Hvilket klassisk spil involverer at stable træklodser, indtil tårnet falder?'},
 'opt1': {'en': 'Kerplunk', 'de': 'Kerplunk', 'es': 'Kerplunk', 'nl': 'Kerplunk', 'sv': 'Kerplunk', 'fr': 'Kerplunk', 'da': 'Kerplunk'},
 'opt2': {'en': 'Jenga', 'de': 'Jenga', 'es': 'Jenga', 'nl': 'Jenga', 'sv': 'Jenga', 'fr': 'Jenga', 'da': 'Jenga'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which yellow spice gives curry its color and is known for anti-inflammatory properties?', 'de': 'Welches gelbe Gewuerz verleiht Curry seine Farbe und ist fuer entzuendungshemmende Eigenschaften bekannt?', 'es': '¿Que especia amarilla le da color al curry y es conocida por sus propiedades antiinflamatorias?', 'nl': 'Welke gele specerij geeft kerrie zijn kleur en staat bekend om ontstekingsremmende eigenschappen?', 'sv': 'Vilken gul krydda ger curry dess farg och ar kand for antiinflammatoriska egenskaper?', 'fr': "Quelle epice jaune donne au curry sa couleur et est connue pour ses proprietes anti-inflammatoires ?", 'da': 'Hvilket gult krydderi giver karry dets farve og er kendt for antiinflammatoriske egenskaber?'},
 'opt1': {'en': 'Saffron', 'de': 'Safran', 'es': 'Azafran', 'nl': 'Saffraan', 'sv': 'Saffran', 'fr': 'Le safran', 'da': 'Safran'},
 'opt2': {'en': 'Turmeric', 'de': 'Kurkuma', 'es': 'Curcuma', 'nl': 'Kurkuma', 'sv': 'Gurkmeja', 'fr': 'Le curcuma', 'da': 'Gurkemeje'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which flower is the national flower of Japan alongside the cherry blossom?', 'de': 'Welche Blume ist neben der Kirschbluete die Nationalblume Japans?', 'es': '¿Que flor es la flor nacional de Japon junto con la flor de cerezo?', 'nl': 'Welke bloem is samen met de kersenbloesem de nationale bloem van Japan?', 'sv': 'Vilken blomma ar Japans nationalblomma tillsammans med korsbarsblomman?', 'fr': "Quelle fleur est la fleur nationale du Japon aux cotes de la fleur de cerisier ?", 'da': 'Hvilken blomst er Japans nationalblomst sammen med kirsebærblomsten?'},
 'opt1': {'en': 'Iris', 'de': 'Iris', 'es': 'Iris', 'nl': 'Iris', 'sv': 'Iris', 'fr': 'Iris', 'da': 'Iris'},
 'opt2': {'en': 'Chrysanthemum', 'de': 'Chrysantheme', 'es': 'Crisantemo', 'nl': 'Chrysant', 'sv': 'Krysantemum', 'fr': 'Chrysantheme', 'da': 'Krysantemum'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. state is home to the Space Needle in Seattle?', 'de': 'In welchem US-Bundesstaat liegt die Space Needle in Seattle?', 'es': '¿En que estado de EE.UU. esta la Space Needle en Seattle?', 'nl': 'In welke Amerikaanse staat ligt de Space Needle in Seattle?', 'sv': 'I vilken amerikansk delstat ligger Space Needle i Seattle?', 'fr': "Dans quel État americain se trouve la Space Needle a Seattle ?", 'da': 'I hvilken amerikansk delstat ligger Space Needle i Seattle?'},
 'opt1': {'en': 'Oregon', 'de': 'Oregon', 'es': 'Oregon', 'nl': 'Oregon', 'sv': 'Oregon', 'fr': "l'Oregon", 'da': 'Oregon'},
 'opt2': {'en': 'Washington', 'de': 'Washington', 'es': 'Washington', 'nl': 'Washington', 'sv': 'Washington', 'fr': 'le Washington', 'da': 'Washington'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which craft involves carving wood into decorative or functional shapes?', 'de': 'Bei welchem Handwerk wird Holz zu dekorativen oder funktionalen Formen geschnitzt?', 'es': '¿Que manualidad implica tallar madera en formas decorativas o funcionales?', 'nl': 'Bij welk handwerk snijd je hout in decoratieve of functionele vormen?', 'sv': 'Vilket hantverk innebar att snida trad till dekorativa eller funktionella former?', 'fr': "Quel art manuel consiste a sculpter le bois en formes decoratives ou fonctionnelles ?", 'da': 'Hvilket håndværk involverer at udskære træ i dekorative eller funktionelle former?'},
 'opt1': {'en': 'Whittling', 'de': 'Schnitzen', 'es': 'Talla en madera', 'nl': 'Houtsnijwerk', 'sv': 'Trasnideri', 'fr': 'Le sculptage au couteau', 'da': 'Træsnitning'},
 'opt2': {'en': 'Woodworking', 'de': 'Holzarbeit', 'es': 'Carpinteria', 'nl': 'Houtbewerking', 'sv': 'Trasnickeri', 'fr': 'Le travail du bois', 'da': 'Træarbejde'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves paddling a small boat through rapids?', 'de': 'Bei welcher Sportart paddelt man ein kleines Boot durch Stromschnellen?', 'es': '¿Que deporte consiste en remar un pequeño bote a traves de rapidos?', 'nl': 'Bij welke sport peddel je een klein bootje door stroomversnellingen?', 'sv': 'Vilken sport innebar att paddla en liten bat genom forsar?', 'fr': "Quel sport consiste a pagayer un petit bateau a travers des rapides ?", 'da': 'Hvilken sport involverer at padle en lille båd gennem stryg?'},
 'opt1': {'en': 'Kayaking', 'de': 'Kajakfahren', 'es': 'Kayak', 'nl': 'Kajakken', 'sv': 'Kajakpaddling', 'fr': 'Le kayak', 'da': 'Kajaksejlads'},
 'opt2': {'en': 'Rowing', 'de': 'Rudern', 'es': 'Remo', 'nl': 'Roeien', 'sv': 'Rodd', 'fr': "L'aviron", 'da': 'Roning'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which item of clothing is worn around the waist to hold up pants?', 'de': 'Welches Kleidungsstueck wird um die Taille getragen, um Hosen zu halten?', 'es': '¿Que prenda se lleva alrededor de la cintura para sostener los pantalones?', 'nl': 'Welk kledingstuk wordt om de taille gedragen om broeken op te houden?', 'sv': 'Vilket klädesplagg baers runt midjan for att halla upp byxor?', 'fr': "Quel vetement se porte autour de la taille pour maintenir le pantalon ?", 'da': 'Hvilken beklædningsgenstand bæres om taljen for at holde bukserne oppe?'},
 'opt1': {'en': 'Belt', 'de': 'Guertel', 'es': 'Cinturon', 'nl': 'Riem', 'sv': 'Bälte', 'fr': 'Ceinture', 'da': 'Bælte'},
 'opt2': {'en': 'Suspenders', 'de': 'Hosentraeger', 'es': 'Tirantes', 'nl': 'Bretels', 'sv': 'Hangslen', 'fr': 'Bretelles', 'da': 'Seler'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic cartoon features a bald-headed boy and his dog Snoopy?', 'de': 'In welchem klassischen Cartoon kommen ein Junge mit kahlem Kopf und sein Hund Snoopy vor?', 'es': '¿Que dibujo animado clasico presenta a un niño calvo y a su perro Snoopy?', 'nl': 'In welke klassieke tekenfilm komen een kaalhoofdige jongen en zijn hond Snoopy voor?', 'sv': 'I vilken klassisk tecknad film forekommer en flintskallig pojke och hans hund Snobben?', 'fr': "Quel dessin anime classique met en scene un garçon chauve et son chien Snoopy ?", 'da': 'Hvilken klassisk tegnefilm har en skaldet dreng og hans hund Snoopy?'},
 'opt1': {'en': 'Peanuts (Charlie Brown)', 'de': 'Die Peanuts (Charlie Brown)', 'es': 'Snoopy y Charlie Brown (Peanuts)', 'nl': 'Peanuts (Charlie Brown)', 'sv': 'Snobben (Charlie Brown)', 'fr': 'Snoopy (Peanuts)', 'da': 'Radiserne (Snoopy)'},
 'opt2': {'en': 'Garfield', 'de': 'Garfield', 'es': 'Garfield', 'nl': 'Garfield', 'sv': 'Garfield', 'fr': 'Garfield', 'da': 'Garfield'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the Dolomite mountains?', 'de': 'In welchem Land liegen die Dolomiten?', 'es': '¿En que pais estan las montañas Dolomitas?', 'nl': 'In welk land liggen de Dolomieten?', 'sv': 'I vilket land ligger Dolomiterna?', 'fr': "Dans quel pays se trouvent les Dolomites ?", 'da': 'Hvilket land huser Dolomitterne?'},
 'opt1': {'en': 'Italy', 'de': 'Italien', 'es': 'Italia', 'nl': 'Italie', 'sv': 'Italien', 'fr': "l'Italie", 'da': 'Italien'},
 'opt2': {'en': 'Austria', 'de': 'Österreich', 'es': 'Austria', 'nl': 'Oostenrijk', 'sv': 'Osterrike', 'fr': "l'Autriche", 'da': 'Østrig'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was a champion of environmental conservation and highway beautification?', 'de': 'Welche First Lady setzte sich fuer Umweltschutz und Autobahnverschoenerung ein?', 'es': '¿Que primera dama fue defensora de la conservacion ambiental y el embellecimiento de carreteras?', 'nl': 'Welke First Lady was voorvechter van natuurbehoud en het verfraaien van snelwegen?', 'sv': 'Vilken First Lady var forespakare for miljobevarande och vagforskönning?', 'fr': "Quelle Première dame etait une championne de la conservation environnementale et de l'embellissement des autoroutes ?", 'da': 'Hvilken førstedame var en fortaler for miljøbevaring og motorvejsforskønnelse?'},
 'opt1': {'en': 'Lady Bird Johnson', 'de': 'Lady Bird Johnson', 'es': 'Lady Bird Johnson', 'nl': 'Lady Bird Johnson', 'sv': 'Lady Bird Johnson', 'fr': 'Lady Bird Johnson', 'da': 'Lady Bird Johnson'},
 'opt2': {'en': 'Eleanor Roosevelt', 'de': 'Eleanor Roosevelt', 'es': 'Eleanor Roosevelt', 'nl': 'Eleanor Roosevelt', 'sv': 'Eleanor Roosevelt', 'fr': 'Eleanor Roosevelt', 'da': 'Eleanor Roosevelt'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Argentinian dish is a grilled assortment of various meats?', 'de': 'Welches argentinische Gericht ist eine gegrillte Auswahl verschiedener Fleischsorten?', 'es': '¿Que plato argentino es un surtido de carnes a la parrilla?', 'nl': 'Welk Argentijns gerecht is een gegrilde assortiment van diverse vleessoorten?', 'sv': 'Vilken argentinsk ratt ar ett grillat urval av olika kottsorter?', 'fr': "Quel plat argentin est un assortiment grille de differentes viandes ?", 'da': 'Hvilken argentinsk ret er et grillet udvalg af forskellige kødtyper?'},
 'opt1': {'en': 'Churrasco', 'de': 'Churrasco', 'es': 'Churrasco', 'nl': 'Churrasco', 'sv': 'Churrasco', 'fr': 'Churrasco', 'da': 'Churrasco'},
 'opt2': {'en': 'Asado', 'de': 'Asado', 'es': 'Asado', 'nl': 'Asado', 'sv': 'Asado', 'fr': 'Asado', 'da': 'Asado'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which hormone regulates the sleep-wake cycle and is produced in darkness?', 'de': 'Welches Hormon reguliert den Schlaf-Wach-Zyklus und wird im Dunkeln produziert?', 'es': '¿Que hormona regula el ciclo de sueño-vigilia y se produce en la oscuridad?', 'nl': 'Welk hormoon reguleert de slaap-waakcyclus en wordt geproduceerd in het donker?', 'sv': 'Vilket hormon reglerar somn-vakenhetscykeln och produceras i morker?', 'fr': "Quelle hormone regule le cycle veille-sommeil et est produite dans l'obscurite ?", 'da': 'Hvilket hormon regulerer søvn-vågencyklussen og produceres i mørke?'},
 'opt1': {'en': 'Cortisol', 'de': 'Cortisol', 'es': 'Cortisol', 'nl': 'Cortisol', 'sv': 'Kortisol', 'fr': 'Cortisol', 'da': 'Kortisol'},
 'opt2': {'en': 'Melatonin', 'de': 'Melatonin', 'es': 'Melatonina', 'nl': 'Melatonine', 'sv': 'Melatonin', 'fr': 'Melatonine', 'da': 'Melatonin'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'The Secret Garden'?", 'de': "Welche Autorin schrieb 'Der geheime Garten'?", 'es': "¿Que autora escribio 'El jardin secreto'?", 'nl': "Welke auteur schreef 'De Geheime Tuin'?", 'sv': "Vilken forfattare skrev 'Hemliga tradgarden'?", 'fr': "Quelle autrice a ecrit 'Le Jardin secret' ?", 'da': "Hvilken forfatter skrev 'Den hemmelige have'?"},
 'opt1': {'en': 'L.M. Montgomery', 'de': 'L.M. Montgomery', 'es': 'L.M. Montgomery', 'nl': 'L.M. Montgomery', 'sv': 'L.M. Montgomery', 'fr': 'L.M. Montgomery', 'da': 'L.M. Montgomery'},
 'opt2': {'en': 'Frances Hodgson Burnett', 'de': 'Frances Hodgson Burnett', 'es': 'Frances Hodgson Burnett', 'nl': 'Frances Hodgson Burnett', 'sv': 'Frances Hodgson Burnett', 'fr': 'Frances Hodgson Burnett', 'da': 'Frances Hodgson Burnett'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Maria von Trapp in the 1965 film 'The Sound of Music'?", 'de': "Welche Schauspielerin spielte Maria von Trapp im Film 'Meine Lieder, meine Traeume' von 1965?", 'es': "¿Que actriz interpreto a Maria von Trapp en la pelicula de 1965 'La novicia rebelde'?", 'nl': "Welke actrice speelde Maria von Trapp in de film 'The Sound of Music' uit 1965?", 'sv': "Vilken skadespelerska spelade Maria von Trapp i 1965 ars film 'Sound of Music'?", 'fr': "Quelle actrice a joue Maria von Trapp dans le film 'La Melodie du bonheur' de 1965 ?", 'da': "Hvilken skuespiller spillede Maria von Trapp i 1965-filmen 'Sound of Music'?"},
 'opt1': {'en': 'Audrey Hepburn', 'de': 'Audrey Hepburn', 'es': 'Audrey Hepburn', 'nl': 'Audrey Hepburn', 'sv': 'Audrey Hepburn', 'fr': 'Audrey Hepburn', 'da': 'Audrey Hepburn'},
 'opt2': {'en': 'Julie Andrews', 'de': 'Julie Andrews', 'es': 'Julie Andrews', 'nl': 'Julie Andrews', 'sv': 'Julie Andrews', 'fr': 'Julie Andrews', 'da': 'Julie Andrews'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which singer released 'Purple Rain' and was known for his eclectic style?", 'de': "Welcher Saenger veroeffentlichte 'Purple Rain' und war fuer seinen exzentrischen Stil bekannt?", 'es': "¿Que cantante lanzo 'Purple Rain' y fue conocido por su estilo eclectico?", 'nl': "Welke zanger bracht 'Purple Rain' uit en stond bekend om zijn eclectische stijl?", 'sv': "Vilken sangare slappte 'Purple Rain' och var kand for sin eklektiska stil?", 'fr': "Quel chanteur a sorti 'Purple Rain' et etait connu pour son style eclectique ?", 'da': "Hvilken sanger udgav 'Purple Rain' og var kendt for sin eklektiske stil?"},
 'opt1': {'en': 'Prince', 'de': 'Prince', 'es': 'Prince', 'nl': 'Prince', 'sv': 'Prince', 'fr': 'Prince', 'da': 'Prince'},
 'opt2': {'en': 'Michael Jackson', 'de': 'Michael Jackson', 'es': 'Michael Jackson', 'nl': 'Michael Jackson', 'sv': 'Michael Jackson', 'fr': 'Michael Jackson', 'da': 'Michael Jackson'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Great Zimbabwe?', 'de': 'In welchem Land liegt die antike Stadt Great Zimbabwe?', 'es': '¿En que pais esta la antigua ciudad de Great Zimbabwe?', 'nl': 'In welk land ligt de oude stad Great Zimbabwe?', 'sv': 'I vilket land ligger den antika staden Great Zimbabwe?', 'fr': "Dans quel pays se trouve l'ancienne cite du Grand Zimbabwe ?", 'da': 'Hvilket land huser den antikke by Great Zimbabwe?'},
 'opt1': {'en': 'Botswana', 'de': 'Botswana', 'es': 'Botsuana', 'nl': 'Botswana', 'sv': 'Botswana', 'fr': 'le Botswana', 'da': 'Botswana'},
 'opt2': {'en': 'Zimbabwe', 'de': 'Simbabwe', 'es': 'Zimbabue', 'nl': 'Zimbabwe', 'sv': 'Zimbabwe', 'fr': 'le Zimbabwe', 'da': 'Zimbabwe'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which small, spiny mammal curls into a ball when threatened, common in the UK?', 'de': 'Welches kleine, stachelige Saeugetier rollt sich bei Bedrohung zu einer Kugel zusammen und ist im UK verbreitet?', 'es': '¿Que pequeño mamifero espinoso se enrolla en una bola cuando es amenazado, comun en el Reino Unido?', 'nl': 'Welk klein, stekelig zoogdier rolt zich op tot een bal bij bedreiging, gebruikelijk in het VK?', 'sv': 'Vilket litet, taggigt daggdjur rullar ihop sig till ett klot vid hot, vanligt i Storbritannien?', 'fr': "Quel petit mammifere epineux s'enroule en boule en cas de menace, courant au Royaume-Uni ?", 'da': 'Hvilket lille, pigget pattedyr ruller sig sammen til en kugle, når det trues, almindeligt i Storbritannien?'},
 'opt1': {'en': 'Porcupine', 'de': 'Stachelschwein', 'es': 'Puerco espin', 'nl': 'Stekelvarken', 'sv': 'Piggsvin', 'fr': 'Porc-epic', 'da': 'Hulepindsvin'},
 'opt2': {'en': 'Hedgehog', 'de': 'Igel', 'es': 'Erizo', 'nl': 'Egel', 'sv': 'Igelkott', 'fr': 'Herisson', 'da': 'Pindsvin'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which German dish is fermented cabbage often served with sausages?', 'de': 'Welches deutsche Gericht ist fermentierter Kohl, oft mit Wuersten serviert?', 'es': '¿Que plato aleman es col fermentada a menudo servida con salchichas?', 'nl': 'Welk Duits gerecht is gefermenteerde kool vaak geserveerd met worsten?', 'sv': 'Vilken tysk ratt ar fermenterad kal ofta serverad med korv?', 'fr': "Quel plat allemand est du chou fermente souvent servi avec des saucisses ?", 'da': 'Hvilken tysk ret er fermenteret kål ofte serveret med pølser?'},
 'opt1': {'en': 'Kimchi', 'de': 'Kimchi', 'es': 'Kimchi', 'nl': 'Kimchi', 'sv': 'Kimchi', 'fr': 'Kimchi', 'da': 'Kimchi'},
 'opt2': {'en': 'Sauerkraut', 'de': 'Sauerkraut', 'es': 'Chucrut', 'nl': 'Zuurkool', 'sv': 'Surkal', 'fr': 'Choucroute', 'da': 'Surkål'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady married President Grover Cleveland in a White House wedding ceremony, the only one held there for a sitting president?', 'de': 'Welche First Lady heiratete Praesident Grover Cleveland in einer Hochzeitszeremonie im Weissen Haus, der einzigen fuer einen amtierenden Praesidenten?', 'es': '¿Que primera dama se caso con el presidente Grover Cleveland en una boda en la Casa Blanca, la unica celebrada alli para un presidente en funciones?', 'nl': 'Welke First Lady trouwde met president Grover Cleveland in een huwelijksceremonie in het Witte Huis, de enige die daar werd gehouden voor een zittende president?', 'sv': 'Vilken First Lady gifte sig med president Grover Cleveland i en brollopsceremoni i Vita huset, den enda som hallits dar for en sittande president?', 'fr': "Quelle Première dame a epouse le president Grover Cleveland lors d'un mariage a la Maison-Blanche, le seul organise la pour un president en exercice ?", 'da': 'Hvilken førstedame giftede sig med præsident Grover Cleveland i en brylluppsceremoni i Det Hvide Hus, den eneste afholdt der for en siddende præsident?'},
 'opt1': {'en': 'Frances Cleveland', 'de': 'Frances Cleveland', 'es': 'Frances Cleveland', 'nl': 'Frances Cleveland', 'sv': 'Frances Cleveland', 'fr': 'Frances Cleveland', 'da': 'Frances Cleveland'},
 'opt2': {'en': 'Julia Tyler', 'de': 'Julia Tyler', 'es': 'Julia Tyler', 'nl': 'Julia Tyler', 'sv': 'Julia Tyler', 'fr': 'Julia Tyler', 'da': 'Julia Tyler'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which artist is known for elongated, thin bronze sculptures of walking figures?', 'de': 'Welcher Kuenstler ist fuer langgezogene, duenne Bronzeskulpturen gehender Figuren bekannt?', 'es': '¿Que artista es conocido por esculturas de bronce alargadas y delgadas de figuras caminando?', 'nl': 'Welke kunstenaar staat bekend om langgerekte, dunne bronzen sculpturen van lopende figuren?', 'sv': 'Vilken konstnar ar kand for avlanga, tunna bronsskulpturer av gaende figurer?', 'fr': "Quel artiste est connu pour ses sculptures en bronze allongees et fines de personnages marchant ?", 'da': 'Hvilken kunstner er kendt for aflange, tynde bronzeskulpturer af gående figurer?'},
 'opt1': {'en': 'Alberto Giacometti', 'de': 'Alberto Giacometti', 'es': 'Alberto Giacometti', 'nl': 'Alberto Giacometti', 'sv': 'Alberto Giacometti', 'fr': 'Alberto Giacometti', 'da': 'Alberto Giacometti'},
 'opt2': {'en': 'Auguste Rodin', 'de': 'Auguste Rodin', 'es': 'Auguste Rodin', 'nl': 'Auguste Rodin', 'sv': 'Auguste Rodin', 'fr': 'Auguste Rodin', 'da': 'Auguste Rodin'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which flowering plant produces the spice saffron from its stigmas?', 'de': 'Welche bluehende Pflanze produziert das Gewuerz Safran aus ihren Narben?', 'es': '¿Que planta con flores produce la especia azafran a partir de sus estigmas?', 'nl': 'Welke bloeiende plant produceert de specerij saffraan uit zijn stampers?', 'sv': 'Vilken blommande vaxt producerar kryddan saffran fran sina martel?', 'fr': "Quelle plante a fleurs produit l'epice safran a partir de ses stigmates ?", 'da': 'Hvilken blomstrende plante producerer krydderiet safran fra sine arr?'},
 'opt1': {'en': 'Crocus', 'de': 'Krokus', 'es': 'Azafran (Crocus)', 'nl': 'Krokus', 'sv': 'Krokus', 'fr': 'Crocus', 'da': 'Krokus'},
 'opt2': {'en': 'Iris', 'de': 'Iris', 'es': 'Iris', 'nl': 'Iris', 'sv': 'Iris', 'fr': 'Iris', 'da': 'Iris'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which musical features the song 'Tomorrow' sung by a young orphan?", 'de': "In welchem Musical kommt das Lied 'Tomorrow' vor, gesungen von einer jungen Waise?", 'es': "¿Que musical presenta la cancion 'Tomorrow' cantada por una joven huerfana?", 'nl': "Welke musical bevat het lied 'Tomorrow' gezongen door een jonge wees?", 'sv': "Vilken musikal innehaller latten 'Tomorrow' sjungen av en ung foraldralos flicka?", 'fr': "Quelle comedie musicale comporte la chanson 'Tomorrow' chantee par une jeune orpheline ?", 'da': "Hvilken musical har sangen 'Tomorrow' sunget af en ung forældreløs?"},
 'opt1': {'en': 'Oliver!', 'de': 'Oliver!', 'es': '¡Oliver!', 'nl': 'Oliver!', 'sv': 'Oliver!', 'fr': 'Oliver!', 'da': 'Oliver!'},
 'opt2': {'en': 'Annie', 'de': 'Annie', 'es': 'Annie', 'nl': 'Annie', 'sv': 'Annie', 'fr': 'Annie', 'da': 'Annie'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country has Helsinki as its capital?', 'de': 'Welches Land hat Helsinki als Hauptstadt?', 'es': '¿Que pais tiene a Helsinki como su capital?', 'nl': 'Welk land heeft Helsinki als hoofdstad?', 'sv': 'Vilket land har Helsingfors som huvudstad?', 'fr': "Quel pays a Helsinki pour capitale ?", 'da': 'Hvilket land har Helsinki som hovedstad?'},
 'opt1': {'en': 'Estonia', 'de': 'Estland', 'es': 'Estonia', 'nl': 'Estland', 'sv': 'Estland', 'fr': "l'Estonie", 'da': 'Estland'},
 'opt2': {'en': 'Finland', 'de': 'Finnland', 'es': 'Finlandia', 'nl': 'Finland', 'sv': 'Finland', 'fr': 'la Finlande', 'da': 'Finland'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which spice is used in traditional Mexican mole sauce alongside chocolate?', 'de': 'Welches Gewuerz wird in traditioneller mexikanischer Mole-Sauce neben Schokolade verwendet?', 'es': '¿Que especia se usa en la salsa mole mexicana tradicional junto al chocolate?', 'nl': 'Welke specerij wordt gebruikt in traditionele Mexicaanse molesaus naast chocolade?', 'sv': 'Vilken krydda anvands i traditionell mexikansk molesas tillsammans med choklad?', 'fr': "Quelle epice est utilisee dans la sauce mole mexicaine traditionnelle aux cotes du chocolat ?", 'da': 'Hvilket krydderi bruges i traditionel mexicansk molesauce ved siden af chokolade?'},
 'opt1': {'en': 'Cumin', 'de': 'Kreuzkuemmel', 'es': 'Comino', 'nl': 'Komijn', 'sv': 'Spiskummin', 'fr': 'Cumin', 'da': 'Spidskommen'},
 'opt2': {'en': 'Chili pepper', 'de': 'Chilischote', 'es': 'Chile', 'nl': 'Chilipeper', 'sv': 'Chilipeppar', 'fr': 'Piment', 'da': 'Chilipeber'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which craft involves creating patterns by dripping or pouring hot wax onto fabric before dyeing?', 'de': 'Bei welchem Handwerk werden Muster erzeugt, indem heisses Wachs auf Stoff getropft oder gegossen wird, bevor gefaerbt wird?', 'es': '¿Que manualidad implica crear patrones goteando o vertiendo cera caliente sobre tela antes de teñirla?', 'nl': 'Bij welk handwerk creëer je patronen door hete was op stof te druppelen of te gieten voordat het geverfd wordt?', 'sv': 'Vilket hantverk innebar att skapa monster genom att droppa eller halla varmt vax pa tyg innan fargning?', 'fr': "Quel art manuel consiste a creer des motifs en versant de la cire chaude sur du tissu avant de le teindre ?", 'da': 'Hvilket håndværk involverer at skabe mønstre ved at dryppe eller hælde varm voks på stof før farvning?'},
 'opt1': {'en': 'Batik', 'de': 'Batik', 'es': 'Batik', 'nl': 'Batik', 'sv': 'Batik', 'fr': 'Le batik', 'da': 'Batik'},
 'opt2': {'en': 'Shibori', 'de': 'Shibori', 'es': 'Shibori', 'nl': 'Shibori', 'sv': 'Shibori', 'fr': 'Shibori', 'da': 'Shibori'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which organ produces insulin to regulate blood sugar?', 'de': 'Welches Organ produziert Insulin, um den Blutzucker zu regulieren?', 'es': '¿Que organo produce insulina para regular el azucar en la sangre?', 'nl': 'Welk orgaan produceert insuline om de bloedsuiker te reguleren?', 'sv': 'Vilket organ producerar insulin for att reglera blodsockret?', 'fr': "Quel organe produit l'insuline pour reguler la glycemie ?", 'da': 'Hvilket organ producerer insulin til at regulere blodsukkeret?'},
 'opt1': {'en': 'Liver', 'de': 'Leber', 'es': 'Higado', 'nl': 'Lever', 'sv': 'Lever', 'fr': 'Foie', 'da': 'Lever'},
 'opt2': {'en': 'Pancreas', 'de': 'Bauchspeicheldruese', 'es': 'Pancreas', 'nl': 'Alvleesklier', 'sv': 'Bukspottkortel', 'fr': 'Pancreas', 'da': 'Bugspytkirtel'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'In rowing, what is the term for the person who steers and coordinates the crew?', 'de': 'Wie nennt man beim Rudern die Person, die steuert und die Mannschaft koordiniert?', 'es': '¿En remo, como se llama a la persona que dirige y coordina a la tripulacion?', 'nl': 'Hoe heet bij het roeien de persoon die stuurt en de bemanning coordineert?', 'sv': 'Vad kallas i rodd personen som styr och koordinerar besattningen?', 'fr': "En aviron, comment appelle-t-on la personne qui dirige et coordonne l'equipage ?", 'da': 'Hvad kaldes i roning den person, der styrer og koordinerer besætningen?'},
 'opt1': {'en': 'Stroke', 'de': 'Schlagmann', 'es': 'Remero de proa', 'nl': 'Slagroeier', 'sv': 'Styrke-roddare', 'fr': 'Chef de nage', 'da': 'Slagroer'},
 'opt2': {'en': 'Coxswain', 'de': 'Steuermann', 'es': 'Timonel', 'nl': 'Stuurman', 'sv': 'Styrman', 'fr': 'Barreur', 'da': 'Styrmand'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Greek island is famous for its blue-domed churches and white buildings?', 'de': 'Welche griechische Insel ist beruehmt fuer ihre blau gekuppelten Kirchen und weissen Gebaeude?', 'es': '¿Que isla griega es famosa por sus iglesias de cupula azul y edificios blancos?', 'nl': 'Welk Grieks eiland staat bekend om zijn blauwe koepelkerken en witte gebouwen?', 'sv': 'Vilken grekisk o ar kand for sina bla kupolkyrkor och vita byggnader?', 'fr': "Quelle ile grecque est celebre pour ses eglises a dome bleu et ses batiments blancs ?", 'da': 'Hvilken græsk ø er berømt for sine blå kuppelkirker og hvide bygninger?'},
 'opt1': {'en': 'Santorini', 'de': 'Santorin', 'es': 'Santorini', 'nl': 'Santorini', 'sv': 'Santorini', 'fr': 'Santorin', 'da': 'Santorini'},
 'opt2': {'en': 'Mykonos', 'de': 'Mykonos', 'es': 'Miconos', 'nl': 'Mykonos', 'sv': 'Mykonos', 'fr': 'Mykonos', 'da': 'Mykonos'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Dorothy in the 1939 film 'The Wizard of Oz'?", 'de': "Welche Schauspielerin spielte Dorothy im Film 'Der Zauberer von Oz' von 1939?", 'es': "¿Que actriz interpreto a Dorothy en la pelicula de 1939 'El Mago de Oz'?", 'nl': "Welke actrice speelde Dorothy in de film 'The Wizard of Oz' uit 1939?", 'sv': "Vilken skadespelerska spelade Dorothy i 1939 ars film 'Trollkarlen fran Oz'?", 'fr': "Quelle actrice a joue Dorothy dans le film 'Le Magicien d'Oz' de 1939 ?", 'da': "Hvilken skuespiller spillede Dorothy i 1939-filmen 'Troldmanden fra Oz'?"},
 'opt1': {'en': 'Judy Garland', 'de': 'Judy Garland', 'es': 'Judy Garland', 'nl': 'Judy Garland', 'sv': 'Judy Garland', 'fr': 'Judy Garland', 'da': 'Judy Garland'},
 'opt2': {'en': 'Shirley Temple', 'de': 'Shirley Temple', 'es': 'Shirley Temple', 'nl': 'Shirley Temple', 'sv': 'Shirley Temple', 'fr': 'Shirley Temple', 'da': 'Shirley Temple'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which classic novel features the character Atticus Finch?', 'de': 'In welchem Klassiker kommt die Figur Atticus Finch vor?', 'es': '¿Que novela clasica presenta al personaje Atticus Finch?', 'nl': 'In welke klassieke roman komt het personage Atticus Finch voor?', 'sv': 'I vilken klassisk roman forekommer karaktaren Atticus Finch?', 'fr': "Quel roman classique met en scene le personnage d'Atticus Finch ?", 'da': 'Hvilken klassisk roman har karakteren Atticus Finch?'},
 'opt1': {'en': 'To Kill a Mockingbird', 'de': 'Wer die Nachtigall stoert', 'es': 'Matar a un ruiseñor', 'nl': 'Spaar de spotvogel', 'sv': 'Skri i natten', 'fr': "Ne tirez pas sur l'oiseau moqueur", 'da': 'Dræb ikke en sangfugl'},
 'opt2': {'en': 'Go Set a Watchman', 'de': 'Gehe hin, stelle einen Waechter', 'es': 'Ve y pon un centinela', 'nl': 'Ga heen, zet een wachter', 'sv': 'Ga, satt ut en vaktare', 'fr': 'Va mettre une sentinelle', 'da': 'Gå hen, sæt en vægter'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which 1980s singer is known for the song 'Like a Virgin'?", 'de': "Welche Saengerin der 1980er ist fuer das Lied 'Like a Virgin' bekannt?", 'es': "¿Que cantante de los 80 es conocida por la cancion 'Like a Virgin'?", 'nl': "Welke zangeres uit de jaren 80 staat bekend om het lied 'Like a Virgin'?", 'sv': "Vilken sangerska fran 1980-talet ar kand for latten 'Like a Virgin'?", 'fr': "Quelle chanteuse des annees 1980 est connue pour la chanson 'Like a Virgin' ?", 'da': "Hvilken 1980’er-sanger er kendt for sangen 'Like a Virgin'?"},
 'opt1': {'en': 'Madonna', 'de': 'Madonna', 'es': 'Madonna', 'nl': 'Madonna', 'sv': 'Madonna', 'fr': 'Madonna', 'da': 'Madonna'},
 'opt2': {'en': 'Cyndi Lauper', 'de': 'Cyndi Lauper', 'es': 'Cyndi Lauper', 'nl': 'Cyndi Lauper', 'sv': 'Cyndi Lauper', 'fr': 'Cyndi Lauper', 'da': 'Cyndi Lauper'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a savory pie filled with sweet potatoes?', 'de': 'Welches suddstaatliche Gericht ist eine herzhafte Torte gefuellt mit Suesskartoffeln?', 'es': '¿Que plato sureño de EE.UU. es un pastel relleno de camote?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een taart gevuld met zoete aardappelen?', 'sv': 'Vilken sydstatlig ratt ar en paj fylld med sotpotatis?', 'fr': "Quel plat du Sud des États-Unis est une tarte fourree aux patates douces ?", 'da': 'Hvilken sydstatsret er en tærte fyldt med søde kartofler?'},
 'opt1': {'en': 'Sweet potato pie', 'de': 'Suesskartoffelpastete', 'es': 'Pastel de camote', 'nl': 'Zoete aardappeltaart', 'sv': 'Sotpotatispaj', 'fr': 'La tarte aux patates douces', 'da': 'Sødkartoffeltærte'},
 'opt2': {'en': 'Pecan pie', 'de': 'Pekannusspastete', 'es': 'Pastel de nuez pecana', 'nl': 'Pecannotentaart', 'sv': 'Pekannötspaj', 'fr': 'La tarte aux noix de pecan', 'da': 'Pekannøddetærte'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which queen ruled the Netherlands for over 50 years in the 19th and 20th centuries?', 'de': 'Welche Koenigin regierte die Niederlande ueber 50 Jahre im 19. und 20. Jahrhundert?', 'es': '¿Que reina goberno los Paises Bajos por mas de 50 años en los siglos XIX y XX?', 'nl': 'Welke koningin regeerde Nederland meer dan 50 jaar in de 19e en 20e eeuw?', 'sv': 'Vilken drottning styrde Nederlanderna i over 50 ar under 1800- och 1900-talen?', 'fr': "Quelle reine a regne sur les Pays-Bas pendant plus de 50 ans aux 19e et 20e siecles ?", 'da': 'Hvilken dronning regerede Nederlandene i over 50 år i det 19. og 20. århundrede?'},
 'opt1': {'en': 'Wilhelmina', 'de': 'Wilhelmina', 'es': 'Guillermina', 'nl': 'Wilhelmina', 'sv': 'Wilhelmina', 'fr': 'Wilhelmine', 'da': 'Wilhelmina'},
 'opt2': {'en': 'Beatrix', 'de': 'Beatrix', 'es': 'Beatriz', 'nl': 'Beatrix', 'sv': 'Beatrix', 'fr': 'Beatrix', 'da': 'Beatrix'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which large aquatic rodent is known for building dams in North America?', 'de': 'Welches grosse Wassernagetier ist fuer den Bau von Daemmen in Nordamerika bekannt?', 'es': '¿Que gran roedor acuatico es conocido por construir presas en Norteamerica?', 'nl': 'Welk groot waterknaagdier staat bekend om het bouwen van dammen in Noord-Amerika?', 'sv': 'Vilken stor vattenlevande gnagare ar kand for att bygga dammar i Nordamerika?', 'fr': "Quel grand rongeur aquatique est connu pour construire des barrages en Amerique du Nord ?", 'da': 'Hvilket stort vandlevende gnaver er kendt for at bygge dæmninger i Nordamerika?'},
 'opt1': {'en': 'Muskrat', 'de': 'Bisamratte', 'es': 'Rata almizclera', 'nl': 'Muskusrat', 'sv': 'Bisamratta', 'fr': 'Rat musque', 'da': 'Bisamrotte'},
 'opt2': {'en': 'Beaver', 'de': 'Biber', 'es': 'Castor', 'nl': 'Bever', 'sv': 'Bavern', 'fr': 'Castor', 'da': 'Bæver'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Tikal?', 'de': 'In welchem Land liegt die antike Stadt Tikal?', 'es': '¿En que pais esta la antigua ciudad de Tikal?', 'nl': 'In welk land ligt de oude stad Tikal?', 'sv': 'I vilket land ligger den antika staden Tikal?', 'fr': "Dans quel pays se trouve l'ancienne cite de Tikal ?", 'da': 'Hvilket land huser den antikke by Tikal?'},
 'opt1': {'en': 'Guatemala', 'de': 'Guatemala', 'es': 'Guatemala', 'nl': 'Guatemala', 'sv': 'Guatemala', 'fr': 'le Guatemala', 'da': 'Guatemala'},
 'opt2': {'en': 'Honduras', 'de': 'Honduras', 'es': 'Honduras', 'nl': 'Honduras', 'sv': 'Honduras', 'fr': 'le Honduras', 'da': 'Honduras'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which type of jacket, popularized by aviators, is made of leather with a fur collar?', 'de': 'Welche Art von Jacke, populaer gemacht von Fliegern, besteht aus Leder mit Pelzkragen?', 'es': '¿Que tipo de chaqueta, popularizada por aviadores, esta hecha de cuero con cuello de piel?', 'nl': 'Welk type jack, populair gemaakt door piloten, is gemaakt van leer met bontkraag?', 'sv': 'Vilken typ av jacka, popularaiserad av flygare, ar gjord av lader med palskrage?', 'fr': "Quel type de veste, popularisee par les aviateurs, est fait de cuir avec un col en fourrure ?", 'da': 'Hvilken type jakke, populariseret af flyvere, er lavet af læder med pelskrave?'},
 'opt1': {'en': 'Denim jacket', 'de': 'Jeansjacke', 'es': 'Chaqueta vaquera', 'nl': 'Spijkerjack', 'sv': 'Jeansjacka', 'fr': 'Veste en jean', 'da': 'Denimjakke'},
 'opt2': {'en': 'Bomber jacket', 'de': 'Bomberjacke', 'es': 'Cazadora bomber', 'nl': 'Bomberjack', 'sv': 'Bomberjacka', 'fr': 'Blouson bombardier', 'da': 'Bomberjakke'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which streaming series follows a group of women running a suburban book club with dark secrets?', 'de': 'Welche Streaming-Serie folgt einer Gruppe von Frauen, die einen Vorstadt-Buchclub mit dunklen Geheimnissen fuehren?', 'es': '¿Que serie de streaming sigue a un grupo de mujeres que dirigen un club de lectura suburbano con oscuros secretos?', 'nl': 'Welke streamingserie volgt een groep vrouwen die een voorstedelijke boekenclub runnen met duistere geheimen?', 'sv': 'Vilken streamingserie foljer en grupp kvinnor som driver en forortsbokklubb med morka hemligheter?', 'fr': "Quelle serie en streaming suit un groupe de femmes dirigeant un club de lecture de banlieue avec de sombres secrets ?", 'da': 'Hvilken streamingserie følger en gruppe kvinder, der driver en forstadsbogklub med mørke hemmeligheder?'},
 'opt1': {'en': 'Big Little Lies', 'de': 'Big Little Lies', 'es': 'Big Little Lies', 'nl': 'Big Little Lies', 'sv': 'Big Little Lies', 'fr': 'Big Little Lies', 'da': 'Big Little Lies'},
 'opt2': {'en': 'Desperate Housewives', 'de': 'Desperate Housewives', 'es': 'Esposas Desesperadas', 'nl': 'Desperate Housewives', 'sv': 'Desperate Housewives', 'fr': 'Desperate Housewives', 'da': 'Desperate Housewives'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which flower is known for its spiral pattern and is a symbol of Fibonacci sequences in nature?', 'de': 'Welche Blume ist fuer ihr Spiralmuster bekannt und ein Symbol fuer Fibonacci-Folgen in der Natur?', 'es': '¿Que flor es conocida por su patron espiral y es un simbolo de las secuencias de Fibonacci en la naturaleza?', 'nl': 'Welke bloem staat bekend om zijn spiraalpatroon en is een symbool van Fibonacci-reeksen in de natuur?', 'sv': 'Vilken blomma ar kand for sitt spiralmonster och ar en symbol for fibonaccital i naturen?', 'fr': "Quelle fleur est connue pour son motif en spirale et est un symbole des suites de Fibonacci dans la nature ?", 'da': 'Hvilken blomst er kendt for sit spiralmønster og er et symbol på Fibonacci-følger i naturen?'},
 'opt1': {'en': 'Sunflower', 'de': 'Sonnenblume', 'es': 'Girasol', 'nl': 'Zonnebloem', 'sv': 'Solros', 'fr': 'Tournesol', 'da': 'Solsikke'},
 'opt2': {'en': 'Daisy', 'de': 'Gaensebluemchen', 'es': 'Margarita', 'nl': 'Madeliefje', 'sv': 'Prastkrage', 'fr': 'Marguerite', 'da': 'Tusindfryd'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Caribbean dish is jerk-seasoned meat, often chicken, grilled over pimento wood?', 'de': 'Welches karibische Gericht ist Jerk-gewuerztes Fleisch, oft Huhn, gegrillt ueber Pimentholz?', 'es': '¿Que plato caribeño es carne sazonada con jerk, a menudo pollo, asada sobre madera de pimienta?', 'nl': 'Welk Caribisch gerecht is jerk-gekruid vlees, vaak kip, gegrild op pimenthout?', 'sv': 'Vilken karibisk ratt ar jerk-kryddat kott, ofta kyckling, grillat over pimentveden?', 'fr': "Quel plat caribeen est de la viande assaisonnee au jerk, souvent du poulet, grillee sur du bois de piment ?", 'da': 'Hvilken caribisk ret er jerk-krydret kød, ofte kylling, grillet over pimentotræ?'},
 'opt1': {'en': 'Jerk chicken', 'de': 'Jerk Chicken', 'es': 'Pollo jerk', 'nl': 'Jerk chicken', 'sv': 'Jerk-kyckling', 'fr': 'Poulet jerk', 'da': 'Jerk-kylling'},
 'opt2': {'en': 'Curry goat', 'de': 'Curry-Ziege', 'es': 'Cabra al curry', 'nl': 'Currygeit', 'sv': 'Curryget', 'fr': 'Curry de chevre', 'da': 'Karrigeder'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which sense is affected by the tongue's taste buds?", 'de': 'Welcher Sinn wird von den Geschmacksknospen der Zunge beeinflusst?', 'es': '¿Que sentido esta afectado por las papilas gustativas de la lengua?', 'nl': 'Welk zintuig wordt beinvloed door de smaakpapillen van de tong?', 'sv': 'Vilket sinne paverkas av tungans smaklokar?', 'fr': "Quel sens est affecte par les papilles gustatives de la langue ?", 'da': 'Hvilken sans påvirkes af tungens smagsløg?'},
 'opt1': {'en': 'Smell', 'de': 'Geruch', 'es': 'Olfato', 'nl': 'Reuk', 'sv': 'Lukt', 'fr': "L'odorat", 'da': 'Lugt'},
 'opt2': {'en': 'Taste', 'de': 'Geschmack', 'es': 'Gusto', 'nl': 'Smaak', 'sv': 'Smak', 'fr': 'Le gout', 'da': 'Smag'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which country is home to the ancient city of Jerusalem's Old City?", 'de': 'In welchem Land liegt die Altstadt von Jerusalem?', 'es': '¿En que pais esta la Ciudad Vieja de Jerusalen?', 'nl': 'In welk land ligt de oude stad van Jeruzalem?', 'sv': 'I vilket land ligger Jerusalems gamla stad?', 'fr': "Dans quel pays se trouve la vieille ville de Jerusalem ?", 'da': 'Hvilket land huser Jerusalems gamle by?'},
 'opt1': {'en': 'Israel', 'de': 'Israel', 'es': 'Israel', 'nl': 'Israel', 'sv': 'Israel', 'fr': 'Israel', 'da': 'Israel'},
 'opt2': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was a former teacher who advocated for education reform?', 'de': 'Welche First Lady war eine ehemalige Lehrerin, die sich fuer Bildungsreformen einsetzte?', 'es': '¿Que primera dama fue una ex maestra que abogo por la reforma educativa?', 'nl': 'Welke First Lady was een voormalige lerares die pleitte voor onderwijshervorming?', 'sv': 'Vilken First Lady var en tidigare larare som forespakade utbildningsreform?', 'fr': "Quelle Première dame etait une ancienne enseignante ayant plaide pour la reforme de l'education ?", 'da': 'Hvilken førstedame var en tidligere lærer, der gik ind for uddannelsesreform?'},
 'opt1': {'en': 'Laura Bush', 'de': 'Laura Bush', 'es': 'Laura Bush', 'nl': 'Laura Bush', 'sv': 'Laura Bush', 'fr': 'Laura Bush', 'da': 'Laura Bush'},
 'opt2': {'en': 'Michelle Obama', 'de': 'Michelle Obama', 'es': 'Michelle Obama', 'nl': 'Michelle Obama', 'sv': 'Michelle Obama', 'fr': 'Michelle Obama', 'da': 'Michelle Obama'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which composer wrote 'Carmina Burana', featuring the famous 'O Fortuna'?", 'de': "Welcher Komponist schrieb 'Carmina Burana' mit dem beruehmten 'O Fortuna'?", 'es': "¿Que compositor escribio 'Carmina Burana', con el famoso 'O Fortuna'?", 'nl': "Welke componist schreef 'Carmina Burana' met het beroemde 'O Fortuna'?", 'sv': "Vilken kompositor skrev 'Carmina Burana', med det beromda 'O Fortuna'?", 'fr': "Quel compositeur a ecrit 'Carmina Burana', avec le celebre 'O Fortuna' ?", 'da': "Hvilken komponist skrev 'Carmina Burana', med det berømte 'O Fortuna'?"},
 'opt1': {'en': 'Carl Orff', 'de': 'Carl Orff', 'es': 'Carl Orff', 'nl': 'Carl Orff', 'sv': 'Carl Orff', 'fr': 'Carl Orff', 'da': 'Carl Orff'},
 'opt2': {'en': 'Gustav Holst', 'de': 'Gustav Holst', 'es': 'Gustav Holst', 'nl': 'Gustav Holst', 'sv': 'Gustav Holst', 'fr': 'Gustav Holst', 'da': 'Gustav Holst'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Vivian Ward in 'Pretty Woman'?", 'de': "Welche Schauspielerin spielte Vivian Ward in 'Pretty Woman'?", 'es': "¿Que actriz interpreto a Vivian Ward en 'Pretty Woman'?", 'nl': "Welke actrice speelde Vivian Ward in 'Pretty Woman'?", 'sv': "Vilken skadespelerska spelade Vivian Ward i 'Pretty Woman'?", 'fr': "Quelle actrice a joue Vivian Ward dans 'Pretty Woman' ?", 'da': "Hvilken skuespiller spillede Vivian Ward i 'Pretty Woman'?"},
 'opt1': {'en': 'Demi Moore', 'de': 'Demi Moore', 'es': 'Demi Moore', 'nl': 'Demi Moore', 'sv': 'Demi Moore', 'fr': 'Demi Moore', 'da': 'Demi Moore'},
 'opt2': {'en': 'Julia Roberts', 'de': 'Julia Roberts', 'es': 'Julia Roberts', 'nl': 'Julia Roberts', 'sv': 'Julia Roberts', 'fr': 'Julia Roberts', 'da': 'Julia Roberts'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Southern U.S. dish is a casserole made from grits and cheese, often with shrimp?', 'de': 'Welches suddstaatliche Gericht ist ein Auflauf aus Grits und Kaese, oft mit Garnelen?', 'es': '¿Que plato sureño de EE.UU. es una cacerola hecha de grits y queso, a menudo con camarones?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een ovenschotel gemaakt van grits en kaas, vaak met garnalen?', 'sv': 'Vilken sydstatlig ratt ar en gryta gjord av grits och ost, ofta med raga?', 'fr': "Quel plat du Sud des États-Unis est un gratin de grits et de fromage, souvent avec des crevettes ?", 'da': 'Hvilken sydstatsret er en gryderet lavet af grits og ost, ofte med rejer?'},
 'opt1': {'en': 'Grits casserole', 'de': 'Grits-Auflauf', 'es': 'Cacerola de grits', 'nl': 'Gritsschotel', 'sv': 'Gritsgryta', 'fr': 'Gratin de grits', 'da': 'Grits-gryderet'},
 'opt2': {'en': 'Shrimp and grits', 'de': 'Shrimp and Grits', 'es': 'Camarones con grits', 'nl': 'Garnalen met grits', 'sv': 'Raga och grits', 'fr': 'Crevettes et grits', 'da': 'Rejer og grits'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which U.S. state is home to the Finger Lakes region?', 'de': 'In welchem US-Bundesstaat liegt die Finger-Lakes-Region?', 'es': '¿En que estado de EE.UU. esta la region de los Finger Lakes?', 'nl': 'In welke Amerikaanse staat ligt de Finger Lakes-regio?', 'sv': 'I vilken amerikansk delstat ligger Finger Lakes-regionen?', 'fr': "Dans quel État americain se trouve la region des Finger Lakes ?", 'da': 'I hvilken amerikansk delstat ligger Finger Lakes-regionen?'},
 'opt1': {'en': 'Pennsylvania', 'de': 'Pennsylvania', 'es': 'Pensilvania', 'nl': 'Pennsylvania', 'sv': 'Pennsylvania', 'fr': 'la Pennsylvanie', 'da': 'Pennsylvania'},
 'opt2': {'en': 'New York', 'de': 'New York', 'es': 'Nueva York', 'nl': 'New York', 'sv': 'New York', 'fr': "l'État de New York", 'da': 'New York'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which large seabird has a distinctive pouched bill for catching fish?', 'de': 'Welcher grosse Seevogel hat einen markanten Beutelschnabel zum Fischfangen?', 'es': '¿Que gran ave marina tiene un pico con bolsa distintivo para atrapar peces?', 'nl': 'Welke grote zeevogel heeft een kenmerkende gezakte snavel om vis te vangen?', 'sv': 'Vilken stor havsfagel har en karakteristisk pungnabb for att fanga fisk?', 'fr': "Quel grand oiseau marin a un bec caracteristique en forme de poche pour attraper les poissons ?", 'da': 'Hvilken stor havfugl har et karakteristisk poset næb til at fange fisk?'},
 'opt1': {'en': 'Cormorant', 'de': 'Kormoran', 'es': 'Cormoran', 'nl': 'Aalscholver', 'sv': 'Skarv', 'fr': 'Cormoran', 'da': 'Skarv'},
 'opt2': {'en': 'Pelican', 'de': 'Pelikan', 'es': 'Pelicano', 'nl': 'Pelikaan', 'sv': 'Pelikan', 'fr': 'Pelican', 'da': 'Pelikan'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Scandinavian tradition involves lighting candles during the darkest days of winter?', 'de': 'Welche skandinavische Tradition beinhaltet das Anzuenden von Kerzen an den dunkelsten Tagen des Winters?', 'es': '¿Que tradicion escandinava implica encender velas durante los dias mas oscuros del invierno?', 'nl': 'Welke Scandinavische traditie omvat het aansteken van kaarsen tijdens de donkerste dagen van de winter?', 'sv': 'Vilken skandinavisk tradition innebar att tanda ljus under vinterns morkaste dagar?', 'fr': "Quelle tradition scandinave implique d'allumer des bougies pendant les jours les plus sombres de l'hiver ?", 'da': 'Hvilken skandinavisk tradition involverer at tænde lys i vinterens mørkeste dage?'},
 'opt1': {'en': 'Advent/St. Lucia tradition', 'de': 'Advent/Lucia-Tradition', 'es': 'Tradicion de Adviento/Santa Lucia', 'nl': 'Advent/Sint-Lucia traditie', 'sv': 'Advent/Luciatraditionen', 'fr': 'La tradition de Sainte-Lucie', 'da': 'Advent/Sankt Lucia-tradition'},
 'opt2': {'en': 'Midsummer', 'de': 'Mittsommer', 'es': 'Solsticio de verano', 'nl': 'Midzomer', 'sv': 'Midsommar', 'fr': 'La Saint-Jean', 'da': 'Sankthans'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'A Wrinkle in Time'?", 'de': "Welche Autorin schrieb 'Falte in der Zeit'?", 'es': "¿Que autora escribio 'Una arruga en el tiempo'?", 'nl': "Welke auteur schreef 'Een Vouw in de Tijd'?", 'sv': "Vilken forfattare skrev 'En sprickai tiden'?", 'fr': "Quelle autrice a ecrit 'Un raccourci dans le temps' ?", 'da': "Hvilken forfatter skrev 'En rynke i tiden'?"},
 'opt1': {'en': 'Judy Blume', 'de': 'Judy Blume', 'es': 'Judy Blume', 'nl': 'Judy Blume', 'sv': 'Judy Blume', 'fr': 'Judy Blume', 'da': 'Judy Blume'},
 'opt2': {'en': "Madeleine L'Engle", 'de': "Madeleine L'Engle", 'es': "Madeleine L'Engle", 'nl': "Madeleine L'Engle", 'sv': "Madeleine L'Engle", 'fr': "Madeleine L'Engle", 'da': "Madeleine L'Engle"}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Petra and the Wadi Rum desert?', 'de': 'In welchem Land liegen die antike Stadt Petra und die Wueste Wadi Rum?', 'es': '¿En que pais estan la antigua ciudad de Petra y el desierto de Wadi Rum?', 'nl': 'In welk land liggen de oude stad Petra en de Wadi Rum-woestijn?', 'sv': 'I vilket land ligger den antika staden Petra och Wadi Rum-oknen?', 'fr': "Dans quel pays se trouvent l'ancienne cite de Petra et le desert du Wadi Rum ?", 'da': 'Hvilket land huser den antikke by Petra og Wadi Rum-ørkenen?'},
 'opt1': {'en': 'Saudi Arabia', 'de': 'Saudi-Arabien', 'es': 'Arabia Saudita', 'nl': 'Saoedi-Arabie', 'sv': 'Saudiarabien', 'fr': 'l\'Arabie saoudite', 'da': 'Saudi-Arabien'},
 'opt2': {'en': 'Jordan', 'de': 'Jordanien', 'es': 'Jordania', 'nl': 'Jordanie', 'sv': 'Jordanien', 'fr': 'la Jordanie', 'da': 'Jordan'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which instrument is central to Irish folk music, a small handheld frame drum?', 'de': 'Welches Instrument ist zentral fuer irische Volksmusik, eine kleine handgehaltene Rahmentrommel?', 'es': '¿Que instrumento es central en la musica folclorica irlandesa, un pequeño tambor de marco de mano?', 'nl': 'Welk instrument is centraal in Ierse volksmuziek, een kleine handbespeelde raamtrom?', 'sv': 'Vilket instrument ar centralt i irlandsk folkmusik, en liten handhallen ramtrumma?', 'fr': "Quel instrument est central dans la musique folklorique irlandaise, un petit tambour sur cadre tenu a la main ?", 'da': 'Hvilket instrument er centralt i irsk folkemusik, en lille håndholdt rammetromme?'},
 'opt1': {'en': 'Bodhran', 'de': 'Bodhran', 'es': 'Bodhran', 'nl': 'Bodhran', 'sv': 'Bodhran', 'fr': 'Bodhran', 'da': 'Bodhran'},
 'opt2': {'en': 'Tambourine', 'de': 'Tamburin', 'es': 'Pandereta', 'nl': 'Tamboerijn', 'sv': 'Tamburin', 'fr': 'Tambourin', 'da': 'Tamburin'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Southern U.S. drink is bourbon mixed with mint and sugar, served at the Kentucky Derby?', 'de': 'Welches suddstaatliche Getraenk ist Bourbon gemischt mit Minze und Zucker, serviert beim Kentucky Derby?', 'es': '¿Que bebida sureña de EE.UU. es bourbon mezclado con menta y azucar, servida en el Derby de Kentucky?', 'nl': 'Welke zuidelijk Amerikaanse drank is bourbon gemengd met munt en suiker, geserveerd bij de Kentucky Derby?', 'sv': 'Vilken sydstatlig dryck ar bourbon blandad med mynta och socker, serverad pa Kentucky Derby?', 'fr': "Quelle boisson du Sud des États-Unis est du bourbon melange a de la menthe et du sucre, servie au Kentucky Derby ?", 'da': 'Hvilken sydstatsdrik er bourbon blandet med mynte og sukker, serveret ved Kentucky Derby?'},
 'opt1': {'en': 'Old Fashioned', 'de': 'Old Fashioned', 'es': 'Old Fashioned', 'nl': 'Old Fashioned', 'sv': 'Old Fashioned', 'fr': 'Old Fashioned', 'da': 'Old Fashioned'},
 'opt2': {'en': 'Mint julep', 'de': 'Mint Julep', 'es': 'Mint julep', 'nl': 'Mint julep', 'sv': 'Mint julep', 'fr': 'Mint julep', 'da': 'Mint julep'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which vitamin is synthesized in the skin from sunlight exposure?', 'de': 'Welches Vitamin wird in der Haut durch Sonnenlichteinstrahlung synthetisiert?', 'es': '¿Que vitamina se sintetiza en la piel por exposicion a la luz solar?', 'nl': 'Welk vitamine wordt in de huid gesynthetiseerd door blootstelling aan zonlicht?', 'sv': 'Vilket vitamin syntetiseras i huden genom exponering for solljus?', 'fr': "Quelle vitamine est synthetisee dans la peau par exposition au soleil ?", 'da': 'Hvilket vitamin syntetiseres i huden ved soleksponering?'},
 'opt1': {'en': 'Vitamin B12', 'de': 'Vitamin B12', 'es': 'Vitamina B12', 'nl': 'Vitamine B12', 'sv': 'Vitamin B12', 'fr': 'Vitamine B12', 'da': 'Vitamin B12'},
 'opt2': {'en': 'Vitamin D', 'de': 'Vitamin D', 'es': 'Vitamina D', 'nl': 'Vitamine D', 'sv': 'Vitamin D', 'fr': 'Vitamine D', 'da': 'Vitamin D'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which film has Sandra Bullock playing an FBI agent who goes undercover in a beauty pageant?', 'de': 'In welchem Film spielt Sandra Bullock eine FBI-Agentin, die undercover an einem Schoenheitswettbewerb teilnimmt?', 'es': '¿En que pelicula Sandra Bullock interpreta a una agente del FBI que va de encubierto en un concurso de belleza?', 'nl': 'In welke film speelt Sandra Bullock een FBI-agent die undercover gaat op een schoonheidswedstrijd?', 'sv': 'I vilken film spelar Sandra Bullock en FBI-agent som gar under tacknamn i en skonhetstavling?', 'fr': "Dans quel film Sandra Bullock joue-t-elle une agente du FBI infiltree dans un concours de beaute ?", 'da': 'I hvilken film spiller Sandra Bullock en FBI-agent, der går undercover i en skønhedskonkurrence?'},
 'opt1': {'en': 'Miss Congeniality', 'de': 'Miss Undercover', 'es': 'Miss Simpatia', 'nl': 'Miss Congeniality', 'sv': 'Miss Undercover', 'fr': 'Miss Detective', 'da': 'Miss Undercover'},
 'opt2': {'en': 'The Proposal', 'de': 'Verlobung auf Umwegen', 'es': 'La propuesta', 'nl': 'The Proposal', 'sv': 'The Proposal', 'fr': 'La Proposition', 'da': 'The Proposal'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Bagan, known for its temples?', 'de': 'In welchem Land liegt die antike Stadt Bagan, bekannt fuer ihre Tempel?', 'es': '¿En que pais esta la antigua ciudad de Bagan, conocida por sus templos?', 'nl': 'In welk land ligt de oude stad Bagan, bekend om zijn tempels?', 'sv': 'I vilket land ligger den antika staden Bagan, kand for sina tempel?', 'fr': "Dans quel pays se trouve l'ancienne cite de Bagan, connue pour ses temples ?", 'da': 'Hvilket land huser den antikke by Bagan, kendt for sine templer?'},
 'opt1': {'en': 'Thailand', 'de': 'Thailand', 'es': 'Tailandia', 'nl': 'Thailand', 'sv': 'Thailand', 'fr': 'la Thailande', 'da': 'Thailand'},
 'opt2': {'en': 'Myanmar', 'de': 'Myanmar', 'es': 'Myanmar', 'nl': 'Myanmar', 'sv': 'Myanmar', 'fr': 'le Myanmar', 'da': 'Myanmar'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of hat, popularized by gangster films, has a wide brim and a creased crown?', 'de': 'Welche Art von Hut, populaer gemacht durch Gangsterfilme, hat eine breite Krempe und eine gefaltete Krone?', 'es': '¿Que tipo de sombrero, popularizado por peliculas de gangsters, tiene ala ancha y copa con pliegue?', 'nl': 'Welk type hoed, populair gemaakt door gangsterfilms, heeft een brede rand en een geplooide bol?', 'sv': 'Vilken typ av hatt, popularaiserad av gangsterfilmer, har brett brätte och veckad puckel?', 'fr': "Quel type de chapeau, popularise par les films de gangsters, a un large bord et une calotte pincee ?", 'da': 'Hvilken type hat, populariseret af gangsterfilm, har en bred skygge og en foldet puld?'},
 'opt1': {'en': 'Fedora', 'de': 'Fedora', 'es': 'Fedora', 'nl': 'Fedora', 'sv': 'Fedora', 'fr': 'Feutre (fedora)', 'da': 'Fedora'},
 'opt2': {'en': 'Bowler hat', 'de': 'Melone (Hut)', 'es': 'Sombrero bombin', 'nl': 'Bolhoed', 'sv': 'Plommonstop', 'fr': 'Chapeau melon', 'da': 'Melonhat'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large deer species is native to North America and known for its antlers and size?', 'de': 'Welche grosse Hirschart ist in Nordamerika heimisch und fuer ihr Geweih und ihre Groesse bekannt?', 'es': '¿Que gran especie de ciervo es nativa de Norteamerica y conocida por su cornamenta y tamaño?', 'nl': 'Welke grote hertensoort is inheems in Noord-Amerika en bekend om zijn gewei en grootte?', 'sv': 'Vilken stor hjortart ar inhemsk i Nordamerika och kand for sina horn och storlek?', 'fr': "Quelle grande espece de cerf est originaire d'Amerique du Nord et connue pour ses bois et sa taille ?", 'da': 'Hvilken stor hjorteart er hjemmehørende i Nordamerika og kendt for sit gevir og størrelse?'},
 'opt1': {'en': 'Moose', 'de': 'Elch', 'es': 'Alce', 'nl': 'Eland', 'sv': 'Alg', 'fr': 'Elan', 'da': 'Elg'},
 'opt2': {'en': 'Elk', 'de': 'Wapiti', 'es': 'Uapiti', 'nl': 'Wapiti', 'sv': 'Wapitihjort', 'fr': 'Wapiti', 'da': 'Wapiti'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which First Lady was known for her 'Vote of Confidence' fashion choices influencing American style?", 'de': "Welche First Lady war bekannt fuer ihre Mode-Entscheidungen, die den amerikanischen Stil beeinflussten?", 'es': "¿Que primera dama fue conocida por sus elecciones de moda que influyeron en el estilo estadounidense?", 'nl': 'Welke First Lady stond bekend om haar modekeuzes die de Amerikaanse stijl beinvloedden?', 'sv': 'Vilken First Lady var kand for sina modeval som paverkade amerikansk stil?', 'fr': "Quelle Première dame etait connue pour ses choix vestimentaires influençant le style americain ?", 'da': 'Hvilken førstedame var kendt for sine modevalg, der påvirkede amerikansk stil?'},
 'opt1': {'en': 'Nancy Reagan', 'de': 'Nancy Reagan', 'es': 'Nancy Reagan', 'nl': 'Nancy Reagan', 'sv': 'Nancy Reagan', 'fr': 'Nancy Reagan', 'da': 'Nancy Reagan'},
 'opt2': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which First Lady's maiden name-inspired rose was named in her honor by horticulturists?", 'de': 'Welche First Lady wurde von Gartenbauern zu Ehren mit einer nach ihr benannten Rosensorte geehrt?', 'es': '¿Que primera dama fue honrada por horticultores con una rosa nombrada en su honor?', 'nl': 'Welke First Lady werd door tuinbouwkundigen geeerd met een naar haar vernoemde roos?', 'sv': 'Vilken First Lady hedrades av tradgardsmastare med en ros uppkallad efter henne?', 'fr': "Quelle Première dame a ete honoree par des horticulteurs avec une rose nommee en son honneur ?", 'da': 'Hvilken førstedame blev hædret af havebrugere med en rose opkaldt efter hende?'},
 'opt1': {'en': 'Nancy Reagan', 'de': 'Nancy Reagan', 'es': 'Nancy Reagan', 'nl': 'Nancy Reagan', 'sv': 'Nancy Reagan', 'fr': 'Nancy Reagan', 'da': 'Nancy Reagan'},
 'opt2': {'en': 'Betty Ford', 'de': 'Betty Ford', 'es': 'Betty Ford', 'nl': 'Betty Ford', 'sv': 'Betty Ford', 'fr': 'Betty Ford', 'da': 'Betty Ford'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which Nigerian author wrote 'Things Fall Apart'?", 'de': "Welcher nigerianische Autor schrieb 'Okonkwo oder Das Alte staerzt'?", 'es': "¿Que autor nigeriano escribio 'Todo se desmorona'?", 'nl': "Welke Nigeriaanse auteur schreef 'Things Fall Apart'?", 'sv': "Vilken nigeriansk forfattare skrev 'Allt gar sonder'?", 'fr': "Quel auteur nigerian a ecrit 'Le Monde s'effondre' ?", 'da': "Hvilken nigeriansk forfatter skrev 'Alt falder fra hinanden'?"},
 'opt1': {'en': 'Wole Soyinka', 'de': 'Wole Soyinka', 'es': 'Wole Soyinka', 'nl': 'Wole Soyinka', 'sv': 'Wole Soyinka', 'fr': 'Wole Soyinka', 'da': 'Wole Soyinka'},
 'opt2': {'en': 'Chinua Achebe', 'de': 'Chinua Achebe', 'es': 'Chinua Achebe', 'nl': 'Chinua Achebe', 'sv': 'Chinua Achebe', 'fr': 'Chinua Achebe', 'da': 'Chinua Achebe'}},

{'category': 'Science', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which vitamin, found in citrus, prevents the disease scurvy in sailors historically?', 'de': 'Welches Vitamin, in Zitrusfruechten enthalten, verhinderte historisch die Krankheit Skorbut bei Seeleuten?', 'es': '¿Que vitamina, presente en citricos, prevenia historicamente el escorbuto en marineros?', 'nl': 'Welk vitamine, aanwezig in citrusvruchten, voorkwam historisch scheurbuik bij zeelieden?', 'sv': 'Vilket vitamin, som finns i citrus, forebyggde historiskt sjukdomen skorbjugg hos sjoman?', 'fr': "Quelle vitamine, presente dans les agrumes, prevenait historiquement le scorbut chez les marins ?", 'da': 'Hvilket vitamin, fundet i citrus, forhindrede historisk sygdommen skørbug hos søfolk?'},
 'opt1': {'en': 'Vitamin C', 'de': 'Vitamin C', 'es': 'Vitamina C', 'nl': 'Vitamine C', 'sv': 'Vitamin C', 'fr': 'Vitamine C', 'da': 'Vitamin C'},
 'opt2': {'en': 'Vitamin A', 'de': 'Vitamin A', 'es': 'Vitamina A', 'nl': 'Vitamine A', 'sv': 'Vitamin A', 'fr': 'Vitamine A', 'da': 'Vitamin A'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which Flemish painter is known for 'The Garden of Earthly Delights'?", 'de': "Welcher flaemische Maler ist fuer 'Der Garten der Lueste' bekannt?", 'es': "¿Que pintor flamenco es conocido por 'El jardin de las delicias'?", 'nl': "Welke Vlaamse schilder staat bekend om 'De Tuin der Lusten'?", 'sv': "Vilken flamlandsk malare ar kand for 'Lustgarden'?", 'fr': "Quel peintre flamand est connu pour 'Le Jardin des delices' ?", 'da': "Hvilken flamsk maler er kendt for 'Lysternes Have'?"},
 'opt1': {'en': 'Pieter Bruegel', 'de': 'Pieter Bruegel', 'es': 'Pieter Brueghel', 'nl': 'Pieter Bruegel', 'sv': 'Pieter Bruegel', 'fr': "Pieter Brueghel", 'da': 'Pieter Bruegel'},
 'opt2': {'en': 'Hieronymus Bosch', 'de': 'Hieronymus Bosch', 'es': 'El Bosco', 'nl': 'Jheronimus Bosch', 'sv': 'Hieronymus Bosch', 'fr': 'Jerome Bosch', 'da': 'Hieronymus Bosch'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which South Asian country is home to Mount Kailash, sacred to multiple religions?', 'de': 'In welchem suedasiatischen Land liegt der Berg Kailash, heilig fuer mehrere Religionen?', 'es': '¿Que pais sudasiatico alberga el Monte Kailash, sagrado para multiples religiones?', 'nl': 'Welk Zuid-Aziatisch land herbergt de berg Kailash, heilig voor meerdere religies?', 'sv': 'Vilket sydasiatiskt land huserar Kailashberget, heligt for flera religioner?', 'fr': "Quel pays d'Asie du Sud abrite le mont Kailash, sacre pour plusieurs religions ?", 'da': 'Hvilket sydasiatisk land huser Kailash-bjerget, helligt for flere religioner?'},
 'opt1': {'en': 'Tibet (China)', 'de': 'Tibet (China)', 'es': 'Tibet (China)', 'nl': 'Tibet (China)', 'sv': 'Tibet (Kina)', 'fr': 'le Tibet (Chine)', 'da': 'Tibet (Kina)'},
 'opt2': {'en': 'Nepal', 'de': 'Nepal', 'es': 'Nepal', 'nl': 'Nepal', 'sv': 'Nepal', 'fr': 'le Nepal', 'da': 'Nepal'}},

{'category': 'Music', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which Finnish composer wrote 'Finlandia'?", 'de': "Welcher finnische Komponist schrieb 'Finlandia'?", 'es': "¿Que compositor finlandes escribio 'Finlandia'?", 'nl': "Welke Finse componist schreef 'Finlandia'?", 'sv': "Vilken finsk kompositor skrev 'Finlandia'?", 'fr': "Quel compositeur finlandais a ecrit 'Finlandia' ?", 'da': "Hvilken finsk komponist skrev 'Finlandia'?"},
 'opt1': {'en': 'Edvard Grieg', 'de': 'Edvard Grieg', 'es': 'Edvard Grieg', 'nl': 'Edvard Grieg', 'sv': 'Edvard Grieg', 'fr': 'Edvard Grieg', 'da': 'Edvard Grieg'},
 'opt2': {'en': 'Jean Sibelius', 'de': 'Jean Sibelius', 'es': 'Jean Sibelius', 'nl': 'Jean Sibelius', 'sv': 'Jean Sibelius', 'fr': 'Jean Sibelius', 'da': 'Jean Sibelius'}},

{'category': 'Food', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Japanese noodle dish is served in a hot broth with pork and soft-boiled egg?', 'de': 'Welches japanische Nudelgericht wird in heisser Bruehe mit Schweinefleisch und weichgekochtem Ei serviert?', 'es': '¿Que plato japones de fideos se sirve en caldo caliente con cerdo y huevo pasado por agua?', 'nl': 'Welk Japans noedelgerecht wordt geserveerd in hete bouillon met varkensvlees en zachtgekookt ei?', 'sv': 'Vilken japansk nudelratt serveras i het buljong med flask och löskokt agg?', 'fr': "Quel plat de nouilles japonais est servi dans un bouillon chaud avec du porc et un oeuf mollet ?", 'da': 'Hvilken japansk nudelret serveres i varm bouillon med svinekød og blødkogt æg?'},
 'opt1': {'en': 'Ramen', 'de': 'Ramen', 'es': 'Ramen', 'nl': 'Ramen', 'sv': 'Ramen', 'fr': 'Ramen', 'da': 'Ramen'},
 'opt2': {'en': 'Udon', 'de': 'Udon', 'es': 'Udon', 'nl': 'Udon', 'sv': 'Udon', 'fr': 'Udon', 'da': 'Udon'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': "In Greek mythology, which winged horse sprang from Medusa's blood?", 'de': 'In der griechischen Mythologie, welches gefluegelte Pferd entsprang dem Blut der Medusa?', 'es': '¿En la mitologia griega, que caballo alado surgio de la sangre de Medusa?', 'nl': 'In de Griekse mythologie, welk gevleugeld paard ontsproot uit het bloed van Medusa?', 'sv': 'I grekisk mytologi, vilken bevingad hast sprang fram ur Medusas blod?', 'fr': "Dans la mythologie grecque, quel cheval aile est ne du sang de Meduse ?", 'da': 'I græsk mytologi, hvilken bevinget hest opstod fra Medusas blod?'},
 'opt1': {'en': 'Pegasus', 'de': 'Pegasus', 'es': 'Pegaso', 'nl': 'Pegasus', 'sv': 'Pegasus', 'fr': 'Pegase', 'da': 'Pegasus'},
 'opt2': {'en': 'Chiron', 'de': 'Chiron', 'es': 'Quiron', 'nl': 'Chiron', 'sv': 'Chiron', 'fr': 'Chiron', 'da': 'Chiron'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which actress played the reclusive, faded silent-film star Norma Desmond in 'Sunset Boulevard'?", 'de': "Welche Schauspielerin spielte den zurueckgezogenen, verblassten Stummfilmstar Norma Desmond in 'Boulevard der Daemmerung'?", 'es': "¿Que actriz interpreto a la reclusa y decadente estrella del cine mudo Norma Desmond en 'El crepusculo de los dioses'?", 'nl': "Welke actrice speelde de teruggetrokken, vervaagde stomme-filmster Norma Desmond in 'Sunset Boulevard'?", 'sv': "Vilken skadespelerska spelade den tillbakadragna, bleknade stumfilmsstjarnan Norma Desmond i 'Sunset Boulevard'?", 'fr': "Quelle actrice a joue Norma Desmond, l'ancienne star du muet recluse, dans 'Boulevard du crepuscule' ?", 'da': "Hvilken skuespiller spillede den tilbagetrukne, falmede stumfilmstjerne Norma Desmond i 'Sunset Boulevard'?"},
 'opt1': {'en': 'Gloria Swanson', 'de': 'Gloria Swanson', 'es': 'Gloria Swanson', 'nl': 'Gloria Swanson', 'sv': 'Gloria Swanson', 'fr': 'Gloria Swanson', 'da': 'Gloria Swanson'},
 'opt2': {'en': 'Bette Davis', 'de': 'Bette Davis', 'es': 'Bette Davis', 'nl': 'Bette Davis', 'sv': 'Bette Davis', 'fr': 'Bette Davis', 'da': 'Bette Davis'}},

{'category': 'Space', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which comet is visible from Earth roughly every 76 years?', 'de': 'Welcher Komet ist von der Erde aus etwa alle 76 Jahre sichtbar?', 'es': '¿Que cometa es visible desde la Tierra aproximadamente cada 76 años?', 'nl': 'Welke komeet is vanaf de aarde ongeveer elke 76 jaar zichtbaar?', 'sv': 'Vilken komet ar synlig fran jorden ungefar vart 76:e ar?', 'fr': "Quelle comete est visible depuis la Terre environ tous les 76 ans ?", 'da': 'Hvilken komet er synlig fra Jorden cirka hvert 76. år?'},
 'opt1': {'en': 'Hale-Bopp', 'de': 'Hale-Bopp', 'es': 'Hale-Bopp', 'nl': 'Hale-Bopp', 'sv': 'Hale-Bopp', 'fr': 'Hale-Bopp', 'da': 'Hale-Bopp'},
 'opt2': {'en': "Halley's Comet", 'de': 'Halleyscher Komet', 'es': 'Cometa Halley', 'nl': 'Komeet van Halley', 'sv': 'Halleys komet', 'fr': 'La comete de Halley', 'da': 'Halleys komet'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which Southeast Asian country's capital is Naypyidaw, a purpose-built city?", 'de': 'Welches suedostasiatische Land hat Naypyidaw, eine eigens erbaute Stadt, als Hauptstadt?', 'es': '¿Que pais del sudeste asiatico tiene como capital a Naipyido, una ciudad construida especialmente?', 'nl': 'Welk Zuidoost-Aziatisch land heeft Naypyidaw, een speciaal gebouwde stad, als hoofdstad?', 'sv': 'Vilket sydostasiatiskt land har Naypyidaw, en specialbyggd stad, som huvudstad?', 'fr': "Quel pays d'Asie du Sud-Est a pour capitale Naypyidaw, une ville construite specialement ?", 'da': 'Hvilket sydøstasiatisk lands hovedstad er Naypyidaw, en specialbygget by?'},
 'opt1': {'en': 'Laos', 'de': 'Laos', 'es': 'Laos', 'nl': 'Laos', 'sv': 'Laos', 'fr': 'le Laos', 'da': 'Laos'},
 'opt2': {'en': 'Myanmar', 'de': 'Myanmar', 'es': 'Myanmar', 'nl': 'Myanmar', 'sv': 'Myanmar', 'fr': 'le Myanmar', 'da': 'Myanmar'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which British designer is known for punk-inspired tartan and bondage-strap fashion?', 'de': 'Welche britische Designerin ist fuer punk-inspirierte Tartan- und Bondage-Riemen-Mode bekannt?', 'es': '¿Que diseñadora britanica es conocida por su moda tartan inspirada en el punk y correas bondage?', 'nl': 'Welke Britse ontwerper staat bekend om punk-geinspireerde tartan- en bondage-riemmode?', 'sv': 'Vilken brittisk designer ar kand for punkinspirerad tartan- och bondagerem-mode?', 'fr': "Quelle creatrice britannique est connue pour sa mode tartan inspiree du punk et ses sangles bondage ?", 'da': 'Hvilken britisk designer er kendt for punk-inspireret tern- og bondage-remmemode?'},
 'opt1': {'en': 'Vivienne Westwood', 'de': 'Vivienne Westwood', 'es': 'Vivienne Westwood', 'nl': 'Vivienne Westwood', 'sv': 'Vivienne Westwood', 'fr': 'Vivienne Westwood', 'da': 'Vivienne Westwood'},
 'opt2': {'en': 'Alexander McQueen', 'de': 'Alexander McQueen', 'es': 'Alexander McQueen', 'nl': 'Alexander McQueen', 'sv': 'Alexander McQueen', 'fr': 'Alexander McQueen', 'da': 'Alexander McQueen'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which English philosopher argued that life without government would be 'nasty, brutish, and short' in 'Leviathan'?", 'de': "Welcher englische Philosoph argumentierte in 'Leviathan', dass das Leben ohne Regierung 'garstig, brutal und kurz' waere?", 'es': "¿Que filosofo ingles argumento en 'Leviatan' que la vida sin gobierno seria 'desagradable, brutal y corta'?", 'nl': "Welke Engelse filosoof betoogde in 'Leviathan' dat het leven zonder overheid 'naar, wreed en kort' zou zijn?", 'sv': "Vilken engelsk filosof argumenterade i 'Leviathan' att livet utan regering skulle vara 'otackt, brutalt och kort'?", 'fr': "Quel philosophe anglais a soutenu dans 'Leviathan' que la vie sans gouvernement serait 'desagreable, brutale et courte' ?", 'da': "Hvilken engelsk filosof argumenterede i 'Leviathan' for, at livet uden regering ville være 'ondt, brutalt og kort'?"},
 'opt1': {'en': 'Thomas Hobbes', 'de': 'Thomas Hobbes', 'es': 'Thomas Hobbes', 'nl': 'Thomas Hobbes', 'sv': 'Thomas Hobbes', 'fr': 'Thomas Hobbes', 'da': 'Thomas Hobbes'},
 'opt2': {'en': 'John Locke', 'de': 'John Locke', 'es': 'John Locke', 'nl': 'John Locke', 'sv': 'John Locke', 'fr': 'John Locke', 'da': 'John Locke'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which quilting pattern uses interlocking rings to resemble a wedding band?', 'de': 'Welches Quiltmuster verwendet ineinandergreifende Ringe, um einem Ehering zu aehneln?', 'es': '¿Que patron de acolchado usa anillos entrelazados para parecer una alianza de boda?', 'nl': 'Welk quiltpatroon gebruikt in elkaar grijpende ringen om op een trouwring te lijken?', 'sv': 'Vilket lapptacksmonster anvander sammanlankade ringar for att likna en vigselring?', 'fr': "Quel motif de patchwork utilise des anneaux entrelaces pour ressembler a une alliance de mariage ?", 'da': 'Hvilket quiltmønster bruger sammenkoblede ringe til at ligne en vielsesring?'},
 'opt1': {'en': 'Double Wedding Ring', 'de': 'Doppelter Ehering', 'es': 'Doble anillo de boda', 'nl': 'Dubbele Trouwring', 'sv': 'Dubbel vigselring', 'fr': 'Double anneau de mariage', 'da': 'Dobbelt vielsesring'},
 'opt2': {'en': 'Dresden Plate', 'de': 'Dresdner Teller', 'es': 'Plato de Dresde', 'nl': 'Dresden Plate', 'sv': 'Dresden Plate', 'fr': 'Assiette de Dresde', 'da': 'Dresden Plate'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which orchid species is famous for its resemblance to a bee, attracting pollinators through mimicry?', 'de': 'Welche Orchideenart ist beruehmt fuer ihre Ähnlichkeit mit einer Biene und zieht Bestaeuber durch Mimikry an?', 'es': '¿Que especie de orquidea es famosa por su parecido con una abeja, atrayendo polinizadores mediante mimetismo?', 'nl': 'Welke orchideeensoort staat bekend om zijn gelijkenis met een bij, waarmee het bestuivers aantrekt door mimicry?', 'sv': 'Vilken orkidear ar beromd for sin likhet med ett bi och lockar pollinerare genom mimikry?', 'fr': "Quelle espece d'orchidee est celebre pour sa ressemblance avec une abeille, attirant les pollinisateurs par mimetisme ?", 'da': 'Hvilken orkidéart er berømt for sin lighed med en bi, tiltrækkende bestøvere gennem mimikry?'},
 'opt1': {'en': 'Moth orchid', 'de': 'Malaienblume', 'es': 'Orquidea mariposa', 'nl': 'Vlinderorchidee', 'sv': 'Fjarilsorkide', 'fr': 'Orchidee papillon', 'da': 'Sommerfugleorkidé'},
 'opt2': {'en': 'Bee orchid', 'de': 'Bienen-Ragwurz', 'es': 'Orquidea abeja', 'nl': 'Bijenorchis', 'sv': 'Biorkide', 'fr': "Ophrys abeille", 'da': 'Bi-gøgeurt'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which architect designed the Chrysler Building's Art Deco spire in New York?", 'de': 'Welcher Architekt entwarf die Art-Deco-Spitze des Chrysler Building in New York?', 'es': '¿Que arquitecto diseño la aguja Art Deco del Chrysler Building en Nueva York?', 'nl': 'Welke architect ontwierp de art-decotoren van het Chrysler Building in New York?', 'sv': 'Vilken arkitekt designade art deco-spiran pa Chrysler Building i New York?', 'fr': "Quel architecte a conçu la fleche Art deco du Chrysler Building a New York ?", 'da': 'Hvilken arkitekt designede Chrysler Building’s art deco-spir i New York?'},
 'opt1': {'en': 'Cass Gilbert', 'de': 'Cass Gilbert', 'es': 'Cass Gilbert', 'nl': 'Cass Gilbert', 'sv': 'Cass Gilbert', 'fr': 'Cass Gilbert', 'da': 'Cass Gilbert'},
 'opt2': {'en': 'William Van Alen', 'de': 'William Van Alen', 'es': 'William Van Alen', 'nl': 'William Van Alen', 'sv': 'William Van Alen', 'fr': 'William Van Alen', 'da': 'William Van Alen'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'In Greek mythology, which hero completed twelve labors as penance?', 'de': 'In der griechischen Mythologie, welcher Held vollbrachte zwoelf Arbeiten als Busse?', 'es': '¿En la mitologia griega, que heroe completo doce trabajos como penitencia?', 'nl': 'In de Griekse mythologie, welke held voltooide twaalf werken als boetedoening?', 'sv': 'I grekisk mytologi, vilken hjalte fullbordade tolv arbeten som bot?', 'fr': "Dans la mythologie grecque, quel heros a accompli douze travaux en penitence ?", 'da': 'I græsk mytologi, hvilken helt fuldførte tolv arbejder som soning?'},
 'opt1': {'en': 'Theseus', 'de': 'Theseus', 'es': 'Teseo', 'nl': 'Theseus', 'sv': 'Theseus', 'fr': 'Thesee', 'da': 'Theseus'},
 'opt2': {'en': 'Hercules', 'de': 'Herkules', 'es': 'Hercules', 'nl': 'Hercules', 'sv': 'Herkules', 'fr': 'Hercule', 'da': 'Herkules'}},
]
