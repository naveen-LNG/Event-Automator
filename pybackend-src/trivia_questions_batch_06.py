# -*- coding: utf-8 -*-
# Trivia Challenge batch #6. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which pink fruit has black seeds and a sweet, juicy interior, popular in summer?', 'de': 'Welche rosa Frucht hat schwarze Kerne und ein suesses, saftiges Inneres, beliebt im Sommer?', 'es': '¿Que fruta rosa tiene semillas negras y un interior dulce y jugoso, popular en verano?', 'nl': 'Welk roze fruit heeft zwarte pitten en een zoet, sappig binnenste, populair in de zomer?', 'sv': 'Vilken rosa frukt har svarta karnor och ett sott, saftigt inre, popular pa sommaren?', 'fr': "Quel fruit rose a des graines noires et un interieur sucre et juteux, populaire en ete ?", 'da': 'Hvilken pink frugt har sorte kerner og et sødt, saftigt indre, populær om sommeren?'},
 'opt1': {'en': 'Cantaloupe', 'de': 'Cantaloupe-Melone', 'es': 'Melon cantalupo', 'nl': 'Cantaloupemeloen', 'sv': 'Cantaloupemelon', 'fr': 'Cantaloup', 'da': 'Cantaloupemelon'},
 'opt2': {'en': 'Watermelon', 'de': 'Wassermelone', 'es': 'Sandia', 'nl': 'Watermeloen', 'sv': 'Vattenmelon', 'fr': 'Pasteque', 'da': 'Vandmelon'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the Kremlin and Red Square?', 'de': 'In welchem Land liegen der Kreml und der Rote Platz?', 'es': '¿En que pais estan el Kremlin y la Plaza Roja?', 'nl': 'In welk land liggen het Kremlin en het Rode Plein?', 'sv': 'I vilket land ligger Kreml och Röda torget?', 'fr': "Dans quel pays se trouvent le Kremlin et la place Rouge ?", 'da': 'Hvilket land huser Kreml og Den Røde Plads?'},
 'opt1': {'en': 'Ukraine', 'de': 'Ukraine', 'es': 'Ucrania', 'nl': 'Oekraine', 'sv': 'Ukraina', 'fr': "l'Ukraine", 'da': 'Ukraine'},
 'opt2': {'en': 'Russia', 'de': 'Russland', 'es': 'Rusia', 'nl': 'Rusland', 'sv': 'Ryssland', 'fr': 'la Russie', 'da': 'Rusland'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a red-haired mermaid who wants to be human?', 'de': 'In welchem Filmklassiker kommt eine rothaarige Meerjungfrau vor, die ein Mensch sein will?', 'es': '¿Que pelicula clasica presenta a una sirena pelirroja que quiere ser humana?', 'nl': 'In welke klassieke film komt een roodharige zeemeermin voor die mens wil worden?', 'sv': 'I vilken klassisk film forekommer en rodharig sjojungfru som vill bli manniska?', 'fr': "Quel film classique met en scene une sirene rousse qui veut devenir humaine ?", 'da': 'Hvilken klassisk film har en rødhåret havfrue, der ønsker at blive menneske?'},
 'opt1': {'en': 'The Little Mermaid', 'de': 'Arielle, die Meerjungfrau', 'es': 'La Sirenita', 'nl': 'De Kleine Zeemeermin', 'sv': 'Den lilla sjöjungfrun', 'fr': 'La Petite Sirene', 'da': 'Den Lille Havfrue'},
 'opt2': {'en': 'Splash', 'de': 'Splash - Eine Jungfrau am Haken', 'es': 'Splash', 'nl': 'Splash', 'sv': 'Splash', 'fr': 'Splash', 'da': 'Splash'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which farm animal is known for gobbling and is served at Thanksgiving?', 'de': 'Welches Nutztier ist fuer sein Kollern bekannt und wird zu Thanksgiving serviert?', 'es': '¿Que animal de granja es conocido por su graznido y se sirve en Accion de Gracias?', 'nl': 'Welk boerderijdier staat bekend om zijn geluid en wordt geserveerd met Thanksgiving?', 'sv': 'Vilket lantbruksdjur ar kant for sitt kluckande och serveras pa Thanksgiving?', 'fr': "Quel animal de ferme est connu pour son glouglou et est servi a Thanksgiving ?", 'da': 'Hvilket landbrugsdyr er kendt for sin skrattende lyd og serveres til Thanksgiving?'},
 'opt1': {'en': 'Turkey', 'de': 'Truthahn', 'es': 'Pavo', 'nl': 'Kalkoen', 'sv': 'Kalkon', 'fr': 'Dinde', 'da': 'Kalkun'},
 'opt2': {'en': 'Chicken', 'de': 'Huhn', 'es': 'Pollo', 'nl': 'Kip', 'sv': 'Kyckling', 'fr': 'Poulet', 'da': 'Kylling'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which yellow condiment is made from soybeans and used in Asian cooking?', 'de': 'Welches gelbe Wuerzmittel wird aus Sojabohnen hergestellt und in der asiatischen Kueche verwendet?', 'es': '¿Que condimento amarillo se hace de semillas de soja y se usa en la cocina asiatica?', 'nl': 'Welk geel smaakmiddel wordt gemaakt van sojabonen en gebruikt in de Aziatische keuken?', 'sv': 'Vilket gult tillbehor gors av sojabonor och anvands i asiatisk mat?', 'fr': "Quel condiment jaune est fait de graines de soja et utilise dans la cuisine asiatique ?", 'da': 'Hvilket gult krydderi er lavet af sojabønner og bruges i asiatisk madlavning?'},
 'opt1': {'en': 'Fish sauce', 'de': 'Fischsauce', 'es': 'Salsa de pescado', 'nl': 'Vissaus', 'sv': 'Fisksas', 'fr': 'Sauce de poisson', 'da': 'Fiskesovs'},
 'opt2': {'en': 'Soy sauce', 'de': 'Sojasauce', 'es': 'Salsa de soja', 'nl': 'Sojasaus', 'sv': 'Sojasas', 'fr': 'Sauce soja', 'da': 'Sojasovs'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to Big Ben and Buckingham Palace?', 'de': 'In welchem Land liegen Big Ben und der Buckingham-Palast?', 'es': '¿En que pais estan el Big Ben y el Palacio de Buckingham?', 'nl': 'In welk land liggen Big Ben en Buckingham Palace?', 'sv': 'I vilket land ligger Big Ben och Buckingham Palace?', 'fr': "Dans quel pays se trouvent Big Ben et le palais de Buckingham ?", 'da': 'Hvilket land huser Big Ben og Buckingham Palace?'},
 'opt1': {'en': 'Ireland', 'de': 'Irland', 'es': 'Irlanda', 'nl': 'Ierland', 'sv': 'Irland', 'fr': "l'Irlande", 'da': 'Irland'},
 'opt2': {'en': 'United Kingdom', 'de': 'Vereinigtes Koenigreich', 'es': 'Reino Unido', 'nl': 'Verenigd Koninkrijk', 'sv': 'Storbritannien', 'fr': 'le Royaume-Uni', 'da': 'Storbritannien'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. president is known for the New Deal during the Great Depression?', 'de': 'Welcher US-Praesident ist fuer den New Deal waehrend der Grossen Depression bekannt?', 'es': '¿Que presidente de EE.UU. es conocido por el New Deal durante la Gran Depresion?', 'nl': 'Welke Amerikaanse president staat bekend om de New Deal tijdens de Grote Depressie?', 'sv': 'Vilken amerikansk president ar kand for New Deal under den stora depressionen?', 'fr': "Quel president americain est connu pour le New Deal pendant la Grande Depression ?", 'da': 'Hvilken amerikansk præsident er kendt for New Deal under Den Store Depression?'},
 'opt1': {'en': 'Herbert Hoover', 'de': 'Herbert Hoover', 'es': 'Herbert Hoover', 'nl': 'Herbert Hoover', 'sv': 'Herbert Hoover', 'fr': 'Herbert Hoover', 'da': 'Herbert Hoover'},
 'opt2': {'en': 'Franklin D. Roosevelt', 'de': 'Franklin D. Roosevelt', 'es': 'Franklin D. Roosevelt', 'nl': 'Franklin D. Roosevelt', 'sv': 'Franklin D. Roosevelt', 'fr': 'Franklin D. Roosevelt', 'da': 'Franklin D. Roosevelt'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves creating pictures using small pieces of colored glass or tile?', 'de': 'Bei welchem Handwerk werden Bilder mit kleinen Stuecken farbigen Glases oder Fliesen erstellt?', 'es': '¿Que manualidad implica crear imagenes usando pequeñas piezas de vidrio o azulejo de color?', 'nl': 'Bij welk handwerk maak je afbeeldingen met kleine stukjes gekleurd glas of tegels?', 'sv': 'Vilket hantverk innebar att skapa bilder med sma bitar av fargat glas eller kakel?', 'fr': "Quel art manuel consiste a creer des images avec de petits morceaux de verre ou de carreaux colores ?", 'da': 'Hvilket håndværk involverer at skabe billeder ved hjælp af små stykker farvet glas eller fliser?'},
 'opt1': {'en': 'Mosaic', 'de': 'Mosaik', 'es': 'Mosaico', 'nl': 'Mozaiek', 'sv': 'Mosaik', 'fr': 'La mosaique', 'da': 'Mosaik'},
 'opt2': {'en': 'Stained glass', 'de': 'Buntglas', 'es': 'Vidriera', 'nl': 'Glas-in-lood', 'sv': 'Glasmalning', 'fr': 'Le vitrail', 'da': 'Glasmaleri'}},

{'category': 'Music', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which brass instrument uses a slide to change pitch?', 'de': 'Welches Blechblasinstrument verwendet einen Zug, um die Tonhoehe zu aendern?', 'es': '¿Que instrumento de metal usa una vara para cambiar el tono?', 'nl': 'Welk koperblazersinstrument gebruikt een schuif om de toonhoogte te veranderen?', 'sv': 'Vilket blasinstrument anvander ett drag for att andra tonhojd?', 'fr': "Quel instrument de cuivre utilise une coulisse pour changer de hauteur ?", 'da': 'Hvilket messinginstrument bruger et træk til at ændre tonehøjde?'},
 'opt1': {'en': 'French horn', 'de': 'Waldhorn', 'es': 'Trompa', 'nl': 'Franse hoorn', 'sv': 'Valthorn', 'fr': 'Cor', 'da': 'Valdhorn'},
 'opt2': {'en': 'Trombone', 'de': 'Posaune', 'es': 'Trombon', 'nl': 'Trombone', 'sv': 'Trombon', 'fr': 'Trombone', 'da': 'Basun'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which season is associated with blooming flowers and warmer weather after winter?', 'de': 'Welche Jahreszeit wird mit bluehenden Blumen und waermerem Wetter nach dem Winter assoziiert?', 'es': '¿Que estacion se asocia con flores en floracion y clima mas calido despues del invierno?', 'nl': 'Welk seizoen wordt geassocieerd met bloeiende bloemen en warmer weer na de winter?', 'sv': 'Vilken arstid forknippas med blommande blommor och varmare vader efter vintern?', 'fr': "Quelle saison est associee aux fleurs qui s'epanouissent et au temps plus chaud apres l'hiver ?", 'da': 'Hvilken årstid forbindes med blomstrende blomster og varmere vejr efter vinter?'},
 'opt1': {'en': 'Spring', 'de': 'Fruehling', 'es': 'Primavera', 'nl': 'Lente', 'sv': 'Var', 'fr': 'Printemps', 'da': 'Forår'},
 'opt2': {'en': 'Summer', 'de': 'Sommer', 'es': 'Verano', 'nl': 'Zomer', 'sv': 'Sommar', 'fr': 'Ete', 'da': 'Sommer'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic film features a spider who saves a pig named Wilbur?', 'de': 'In welchem Filmklassiker rettet eine Spinne ein Schwein namens Wilbur?', 'es': '¿Que pelicula clasica presenta a una araña que salva a un cerdo llamado Wilbur?', 'nl': 'In welke klassieke film redt een spin een varken genaamd Wilbur?', 'sv': 'I vilken klassisk film raddar en spindel en gris vid namn Wilbur?', 'fr': "Quel film classique met en scene une araignee qui sauve un cochon nomme Wilbur ?", 'da': 'Hvilken klassisk film har en edderkop, der redder en gris ved navn Wilbur?'},
 'opt1': {'en': 'Babe', 'de': 'Schweinchen Babe in der grossen Stadt', 'es': 'Babe, el cerdito valiente', 'nl': 'Babe', 'sv': 'Babe - Den modiga lilla grisen', 'fr': 'Babe, le cochon devenu berger', 'da': 'Babe - Den Frejdige Gris'},
 'opt2': {'en': "Charlotte's Web", 'de': 'Charlottes Netz', 'es': 'La Telaraña de Carlota', 'nl': "Charlotte's Web", 'sv': 'Charlottes nat', 'fr': 'Le Petit Monde de Charlotte', 'da': 'Charlottes Web'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': "Which U.S. state is known as the 'Sunflower State'?", 'de': "Welcher US-Bundesstaat ist als 'Sunflower State' bekannt?", 'es': "¿Que estado de EE.UU. es conocido como el 'Estado del Girasol'?", 'nl': "Welke Amerikaanse staat staat bekend als de 'Sunflower State'?", 'sv': "Vilken amerikansk delstat ar kand som 'Solrosstaten'?", 'fr': "Quel État americain est connu comme le 'Sunflower State' (État du tournesol) ?", 'da': "Hvilken amerikansk delstat er kendt som 'Sunflower State'?"},
 'opt1': {'en': 'Nebraska', 'de': 'Nebraska', 'es': 'Nebraska', 'nl': 'Nebraska', 'sv': 'Nebraska', 'fr': 'le Nebraska', 'da': 'Nebraska'},
 'opt2': {'en': 'Kansas', 'de': 'Kansas', 'es': 'Kansas', 'nl': 'Kansas', 'sv': 'Kansas', 'fr': 'le Kansas', 'da': 'Kansas'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which white vegetable is often mashed and served as a side dish?', 'de': 'Welches weisse Gemuese wird oft pueriert und als Beilage serviert?', 'es': '¿Que verdura blanca se hace pure a menudo y se sirve como guarnicion?', 'nl': 'Welke witte groente wordt vaak gepureerd en als bijgerecht geserveerd?', 'sv': 'Vilken vit gronsak mosas ofta och serveras som tillbehor?', 'fr': "Quel legume blanc est souvent reduit en puree et servi en accompagnement ?", 'da': 'Hvilken hvid grøntsag mostes ofte og serveres som tilbehør?'},
 'opt1': {'en': 'Potato', 'de': 'Kartoffel', 'es': 'Papa', 'nl': 'Aardappel', 'sv': 'Potatis', 'fr': 'La pomme de terre', 'da': 'Kartoffel'},
 'opt2': {'en': 'Turnip', 'de': 'Rube', 'es': 'Nabo', 'nl': 'Raap', 'sv': 'Rova', 'fr': 'Le navet', 'da': 'Roe'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves throwing darts at a circular target?', 'de': 'Bei welcher Sportart wirft man Pfeile auf eine kreisfoermige Zielscheibe?', 'es': '¿Que deporte consiste en lanzar dardos a un blanco circular?', 'nl': 'Bij welke sport gooi je pijltjes naar een cirkelvormig doel?', 'sv': 'Vilken sport innebar att kasta pilar mot en cirkelformad tavla?', 'fr': "Quel sport consiste a lancer des flechettes sur une cible circulaire ?", 'da': 'Hvilken sport involverer at kaste pile mod en cirkulær skive?'},
 'opt1': {'en': 'Darts', 'de': 'Darts', 'es': 'Dardos', 'nl': 'Darten', 'sv': 'Pilkastning', 'fr': 'Les flechettes', 'da': 'Dart'},
 'opt2': {'en': 'Archery', 'de': 'Bogenschiessen', 'es': 'Tiro con arco', 'nl': 'Boogschieten', 'sv': 'Bagskytte', 'fr': "Le tir a l'arc", 'da': 'Bueskydning'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which pet bird is known for its colorful feathers and ability to mimic speech?', 'de': 'Welcher Haustiervogel ist fuer seine bunten Federn und die Faehigkeit bekannt, Sprache nachzuahmen?', 'es': '¿Que ave mascota es conocida por sus plumas coloridas y su capacidad de imitar el habla?', 'nl': 'Welke huisdiervogel staat bekend om zijn kleurrijke veren en het vermogen om spraak na te bootsen?', 'sv': 'Vilken sallskapsfagel ar kand for sina fargglada fjadrar och formaga att imitera tal?', 'fr': "Quel oiseau de compagnie est connu pour ses plumes colorees et sa capacite a imiter la parole ?", 'da': 'Hvilken kæledyrsfugl er kendt for sine farverige fjer og evne til at efterligne tale?'},
 'opt1': {'en': 'Parrot', 'de': 'Papagei', 'es': 'Loro', 'nl': 'Papegaai', 'sv': 'Papegoja', 'fr': 'Perroquet', 'da': 'Papegøje'},
 'opt2': {'en': 'Canary', 'de': 'Kanarienvogel', 'es': 'Canario', 'nl': 'Kanarie', 'sv': 'Kanariefagel', 'fr': 'Canari', 'da': 'Kanariefugl'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which holiday involves searching for hidden eggs in the spring?', 'de': 'Bei welchem Feiertag im Fruehling sucht man versteckte Eier?', 'es': '¿Que festividad de primavera implica buscar huevos escondidos?', 'nl': 'Bij welke feestdag in de lente zoek je naar verstopte eieren?', 'sv': 'Vilken varhogtid innebar att leta efter gomda agg?', 'fr': "Quelle fete de printemps implique de chercher des oeufs caches ?", 'da': 'Hvilken forårshelligdag involverer at lede efter skjulte æg?'},
 'opt1': {'en': 'May Day', 'de': 'Maifeiertag', 'es': 'Dia de mayo', 'nl': 'Meifeest', 'sv': 'Valborg', 'fr': 'Le 1er mai', 'da': 'Majdag'},
 'opt2': {'en': 'Easter', 'de': 'Ostern', 'es': 'Pascua', 'nl': 'Pasen', 'sv': 'Pask', 'fr': 'Pâques', 'da': 'Påske'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which continent is home to the Amazon Rainforest?', 'de': 'Auf welchem Kontinent liegt der Amazonas-Regenwald?', 'es': '¿En que continente esta la selva amazonica?', 'nl': 'Op welk continent ligt het Amazoneregenwoud?', 'sv': 'Pa vilken kontinent ligger Amazonas regnskog?', 'fr': "Sur quel continent se trouve la foret amazonienne ?", 'da': 'Hvilket kontinent huser Amazonas regnskov?'},
 'opt1': {'en': 'South America', 'de': 'Suedamerika', 'es': 'Sudamerica', 'nl': 'Zuid-Amerika', 'sv': 'Sydamerika', 'fr': "l'Amerique du Sud", 'da': 'Sydamerika'},
 'opt2': {'en': 'Africa', 'de': 'Afrika', 'es': 'Africa', 'nl': 'Afrika', 'sv': 'Afrika', 'fr': "l'Afrique", 'da': 'Afrika'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which author wrote 'Green Eggs and Ham'?", 'de': "Welcher Autor schrieb 'Gruene Eier mit Schinken'?", 'es': "¿Que autor escribio 'Huevos verdes con jamon'?", 'nl': "Welke auteur schreef 'Groene Eieren met Ham'?", 'sv': "Vilken forfattare skrev 'Grona agg och skinka'?", 'fr': "Quel auteur a ecrit 'Les Oeufs verts au jambon' ?", 'da': "Hvilken forfatter skrev 'Grønne æg og skinke'?"},
 'opt1': {'en': 'Dr. Seuss', 'de': 'Dr. Seuss', 'es': 'Dr. Seuss', 'nl': 'Dr. Seuss', 'sv': 'Dr. Seuss', 'fr': 'Dr Seuss', 'da': 'Dr. Seuss'},
 'opt2': {'en': 'Shel Silverstein', 'de': 'Shel Silverstein', 'es': 'Shel Silverstein', 'nl': 'Shel Silverstein', 'sv': 'Shel Silverstein', 'fr': 'Shel Silverstein', 'da': 'Shel Silverstein'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic toy is a wooden top that spins on a string?', 'de': 'Welches klassische Spielzeug ist ein Holzkreisel, der an einer Schnur dreht?', 'es': '¿Que juguete clasico es un trompo de madera que gira con una cuerda?', 'nl': 'Welk klassiek speelgoed is een houten tol die draait op een touwtje?', 'sv': 'Vilken klassisk leksak ar en trasnurra som snurrar pa ett snore?', 'fr': "Quel jouet classique est une toupie en bois qui tourne sur une ficelle ?", 'da': 'Hvilket klassisk legetøj er en trætop, der snurrer på en snor?'},
 'opt1': {'en': 'Yo-yo', 'de': 'Jo-Jo', 'es': 'Yoyo', 'nl': 'Jojo', 'sv': 'Jojo', 'fr': 'Yo-yo', 'da': 'Jo-jo'},
 'opt2': {'en': 'Spinning top', 'de': 'Kreisel', 'es': 'Trompo', 'nl': 'Tol', 'sv': 'Snurra', 'fr': 'Toupie', 'da': 'Snurretop'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which brown spice is used in gingerbread and pumpkin spice blends?', 'de': 'Welches braune Gewuerz wird in Lebkuchen und Kuerbisgewuerzmischungen verwendet?', 'es': '¿Que especia marron se usa en el pan de jengibre y las mezclas de especias de calabaza?', 'nl': 'Welke bruine specerij wordt gebruikt in peperkoek en pompoenspecerijmengsels?', 'sv': 'Vilken brun krydda anvands i pepparkakor och pumpakryddblandningar?', 'fr': "Quelle epice brune est utilisee dans le pain d'epices et les melanges d'epices a citrouille ?", 'da': 'Hvilket brunt krydderi bruges i honningkage og græskarkrydderiblandinger?'},
 'opt1': {'en': 'Cinnamon', 'de': 'Zimt', 'es': 'Canela', 'nl': 'Kaneel', 'sv': 'Kanel', 'fr': 'La cannelle', 'da': 'Kanel'},
 'opt2': {'en': 'Nutmeg', 'de': 'Muskatnuss', 'es': 'Nuez moscada', 'nl': 'Nootmuskaat', 'sv': 'Muskot', 'fr': 'La muscade', 'da': 'Muskatnød'}},

{'category': 'History', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. holiday honors military veterans in November?', 'de': 'Welcher US-Feiertag im November ehrt Militaerveteranen?', 'es': '¿Que festividad de EE.UU. en noviembre honra a los veteranos militares?', 'nl': 'Welke Amerikaanse feestdag in november eert militaire veteranen?', 'sv': 'Vilken amerikansk hogtid i november hedrar militarveteraner?', 'fr': "Quelle fete americaine de novembre honore les veterans militaires ?", 'da': 'Hvilken amerikansk helligdag i november hædrer militærveteraner?'},
 'opt1': {'en': 'Veterans Day', 'de': 'Veterans Day', 'es': 'Dia de los Veteranos', 'nl': 'Veterans Day', 'sv': 'Veterandagen', 'fr': 'Le jour des Anciens combattants', 'da': 'Veterandag'},
 'opt2': {'en': 'Memorial Day', 'de': 'Memorial Day', 'es': 'Dia de los Caidos', 'nl': 'Memorial Day', 'sv': 'Memorial Day', 'fr': 'Le Memorial Day', 'da': 'Memorial Day'}},

{'category': 'Music', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which decade is associated with the British Invasion led by bands like the Beatles?', 'de': 'Mit welchem Jahrzehnt wird die British Invasion, angefuehrt von Bands wie den Beatles, assoziiert?', 'es': '¿Con que decada se asocia la Invasion Britanica liderada por bandas como los Beatles?', 'nl': 'Met welk decennium wordt de British Invasion, geleid door bands als de Beatles, geassocieerd?', 'sv': 'Vilket artionde forknippas med den brittiska invasionen ledd av band som Beatles?', 'fr': "A quelle decennie associe-t-on l'invasion britannique menee par des groupes comme les Beatles ?", 'da': 'Hvilket årti forbindes med den britiske invasion ledet af bands som Beatles?'},
 'opt1': {'en': '1950s', 'de': '1950er', 'es': 'los años 50', 'nl': 'jaren 50', 'sv': '1950-talet', 'fr': 'les annees 1950', 'da': '1950’erne'},
 'opt2': {'en': '1960s', 'de': '1960er', 'es': 'los años 60', 'nl': 'jaren 60', 'sv': '1960-talet', 'fr': 'les annees 1960', 'da': '1960’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient Oracle of Delphi?', 'de': 'In welchem Land liegt das antike Orakel von Delphi?', 'es': '¿En que pais esta el antiguo Oraculo de Delfos?', 'nl': 'In welk land ligt het oude Orakel van Delphi?', 'sv': 'I vilket land ligger det antika oraklet i Delfi?', 'fr': "Dans quel pays se trouve l'ancien oracle de Delphes ?", 'da': 'Hvilket land huser det antikke orakel i Delfi?'},
 'opt1': {'en': 'Greece', 'de': 'Griechenland', 'es': 'Grecia', 'nl': 'Griekenland', 'sv': 'Grekland', 'fr': 'la Grece', 'da': 'Grækenland'},
 'opt2': {'en': 'Italy', 'de': 'Italien', 'es': 'Italia', 'nl': 'Italie', 'sv': 'Italien', 'fr': "l'Italie", 'da': 'Italien'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which farm animal produces wool for sweaters and blankets?', 'de': 'Welches Nutztier produziert Wolle fuer Pullover und Decken?', 'es': '¿Que animal de granja produce lana para sueteres y mantas?', 'nl': 'Welk boerderijdier produceert wol voor truien en dekens?', 'sv': 'Vilket lantbruksdjur producerar ull for troejor och filtar?', 'fr': "Quel animal de ferme produit de la laine pour les pulls et les couvertures ?", 'da': 'Hvilket landbrugsdyr producerer uld til sweatre og tæpper?'},
 'opt1': {'en': 'Alpaca', 'de': 'Alpaka', 'es': 'Alpaca', 'nl': 'Alpaca', 'sv': 'Alpacka', 'fr': 'Alpaga', 'da': 'Alpaka'},
 'opt2': {'en': 'Sheep', 'de': 'Schaf', 'es': 'Oveja', 'nl': 'Schaap', 'sv': 'Far', 'fr': 'Mouton', 'da': 'Får'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which flower is known for its trumpet shape and comes in a wide range of vivid colors, popular in spring gardens?', 'de': 'Welche Blume ist fuer ihre Trompetenform bekannt und kommt in vielen leuchtenden Farben vor, beliebt in Fruehlingsgaerten?', 'es': '¿Que flor es conocida por su forma de trompeta y viene en una amplia gama de colores vivos, popular en jardines primaverales?', 'nl': 'Welke bloem staat bekend om zijn trompetvorm en komt voor in een breed scala aan levendige kleuren, populair in lentetuinen?', 'sv': 'Vilken blomma ar kand for sin trumpetform och kommer i ett brett spektrum av levande farger, popular i vartradgardar?', 'fr': "Quelle fleur est connue pour sa forme de trompette et se decline en une large gamme de couleurs vives, populaire dans les jardins printaniers ?", 'da': 'Hvilken blomst er kendt for sin trompetform og fås i mange levende farver, populær i forårshaver?'},
 'opt1': {'en': 'Iris', 'de': 'Iris', 'es': 'Iris', 'nl': 'Iris', 'sv': 'Iris', 'fr': 'Iris', 'da': 'Iris'},
 'opt2': {'en': 'Daffodil', 'de': 'Osterglocke', 'es': 'Narciso', 'nl': 'Narcis', 'sv': 'Pasklilja', 'fr': 'Jonquille', 'da': 'Pinselilje'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic card game involves matching pairs of cards face down?', 'de': 'Bei welchem klassischen Kartenspiel sucht man Paare von umgedrehten Karten?', 'es': '¿Que juego de cartas clasico consiste en emparejar cartas boca abajo?', 'nl': 'Bij welk klassiek kaartspel zoek je bijpassende paren kaarten die omgekeerd liggen?', 'sv': 'Vilket klassiskt kortspel innebar att matcha par av kort som ligger med baksidan upp?', 'fr': "Quel jeu de cartes classique consiste a associer des paires de cartes face cachee ?", 'da': 'Hvilket klassisk kortspil involverer at matche par af kort, der ligger med bagsiden opad?'},
 'opt1': {'en': 'Memory (Concentration)', 'de': 'Memory', 'es': 'Memoria (Concentracion)', 'nl': 'Memory', 'sv': 'Memory', 'fr': 'Memory', 'da': 'Memory'},
 'opt2': {'en': 'Go Fish', 'de': 'Schwarzer Peter', 'es': 'Vete a pescar', 'nl': 'Vissen', 'sv': 'Fiske', 'fr': 'Le Pouilleux', 'da': 'Sorteper'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which green spice is used to make traditional Italian pesto?', 'de': 'Welches gruene Gewuerz wird verwendet, um traditionelles italienisches Pesto herzustellen?', 'es': '¿Que hierba verde se usa para hacer el pesto italiano tradicional?', 'nl': 'Welk groen kruid wordt gebruikt om traditionele Italiaanse pesto te maken?', 'sv': 'Vilken gron krydda anvands for att gora traditionell italiensk pesto?', 'fr': "Quelle herbe verte est utilisee pour faire le pesto italien traditionnel ?", 'da': 'Hvilken grøn urt bruges til at lave traditionel italiensk pesto?'},
 'opt1': {'en': 'Basil', 'de': 'Basilikum', 'es': 'Albahaca', 'nl': 'Basilicum', 'sv': 'Basilika', 'fr': 'Le basilic', 'da': 'Basilikum'},
 'opt2': {'en': 'Oregano', 'de': 'Oregano', 'es': 'Oregano', 'nl': 'Oregano', 'sv': 'Oregano', 'fr': "L'origan", 'da': 'Oregano'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which flower is traditionally given as a corsage at prom?', 'de': 'Welche Blume wird traditionell als Corsage zum Abschlussball geschenkt?', 'es': '¿Que flor se regala tradicionalmente como ramillete en el baile de graduacion?', 'nl': 'Welke bloem wordt traditioneel gegeven als corsage op het schoolbal?', 'sv': 'Vilken blomma ges traditionellt som en corsage pa balen?', 'fr': "Quelle fleur est traditionnellement offerte en corsage au bal de promo ?", 'da': 'Hvilken blomst gives traditionelt som en corsage til studenterballet?'},
 'opt1': {'en': 'Orchid', 'de': 'Orchidee', 'es': 'Orquidea', 'nl': 'Orchidee', 'sv': 'Orkide', 'fr': 'Orchidee', 'da': 'Orkide'},
 'opt2': {'en': 'Rose', 'de': 'Rose', 'es': 'Rosa', 'nl': 'Roos', 'sv': 'Ros', 'fr': 'Rose', 'da': 'Rose'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Alamo in San Antonio?', 'de': 'In welchem US-Bundesstaat liegt das Alamo in San Antonio?', 'es': '¿En que estado de EE.UU. esta el Alamo en San Antonio?', 'nl': 'In welke Amerikaanse staat ligt de Alamo in San Antonio?', 'sv': 'I vilken amerikansk delstat ligger Alamo i San Antonio?', 'fr': "Dans quel État americain se trouve l'Alamo a San Antonio ?", 'da': 'I hvilken amerikansk delstat ligger Alamo i San Antonio?'},
 'opt1': {'en': 'Texas', 'de': 'Texas', 'es': 'Texas', 'nl': 'Texas', 'sv': 'Texas', 'fr': 'le Texas', 'da': 'Texas'},
 'opt2': {'en': 'New Mexico', 'de': 'New Mexico', 'es': 'Nuevo Mexico', 'nl': 'New Mexico', 'sv': 'New Mexico', 'fr': 'le Nouveau-Mexique', 'da': 'New Mexico'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which craft involves sewing beads onto fabric to create designs?', 'de': 'Bei welchem Handwerk werden Perlen auf Stoff genaeht, um Muster zu erstellen?', 'es': '¿Que manualidad implica coser cuentas en tela para crear diseños?', 'nl': 'Bij welk handwerk naai je kralen op stof om patronen te maken?', 'sv': 'Vilket hantverk innebar att sy palror pa tyg for att skapa monster?', 'fr': "Quel art manuel consiste a coudre des perles sur du tissu pour creer des motifs ?", 'da': 'Hvilket håndværk involverer at sy perler på stof for at skabe mønstre?'},
 'opt1': {'en': 'Embroidery', 'de': 'Stickerei', 'es': 'Bordado', 'nl': 'Borduurwerk', 'sv': 'Broderi', 'fr': 'Broderie', 'da': 'Broderi'},
 'opt2': {'en': 'Beadwork', 'de': 'Perlenstickerei', 'es': 'Abalorios', 'nl': 'Kralenwerk', 'sv': 'Palrbroderi', 'fr': 'Le perlage', 'da': 'Perlebroderi'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport uses a small ball hit back and forth over a table with paddles?', 'de': 'Bei welcher Sportart wird ein kleiner Ball mit Schlaegern ueber einen Tisch hin und her geschlagen?', 'es': '¿Que deporte usa una pequeña pelota golpeada de un lado a otro sobre una mesa con paletas?', 'nl': 'Bij welke sport wordt een kleine bal met batjes over een tafel heen en weer geslagen?', 'sv': 'Vilken sport anvander en liten boll som slas fram och tillbaka over ett bord med racketar?', 'fr': "Quel sport utilise une petite balle frappee d'avant en arriere sur une table avec des raquettes ?", 'da': 'Hvilken sport bruger en lille bold, der slås frem og tilbage over et bord med bat?'},
 'opt1': {'en': 'Table tennis', 'de': 'Tischtennis', 'es': 'Tenis de mesa', 'nl': 'Tafeltennis', 'sv': 'Bordtennis', 'fr': 'Le tennis de table', 'da': 'Bordtennis'},
 'opt2': {'en': 'Badminton', 'de': 'Badminton', 'es': 'Badminton', 'nl': 'Badminton', 'sv': 'Badminton', 'fr': 'Le badminton', 'da': 'Badminton'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which item of clothing covers the legs and is worn under a skirt in cold weather?', 'de': 'Welches Kleidungsstueck bedeckt die Beine und wird bei kaltem Wetter unter einem Rock getragen?', 'es': '¿Que prenda cubre las piernas y se lleva bajo una falda en clima frio?', 'nl': 'Welk kledingstuk bedekt de benen en wordt onder een rok gedragen bij koud weer?', 'sv': 'Vilket klädesplagg tacker benen och baers under en kjol i kallt vader?', 'fr': "Quel vetement couvre les jambes et se porte sous une jupe par temps froid ?", 'da': 'Hvilken beklædningsgenstand dækker benene og bæres under en nederdel i koldt vejr?'},
 'opt1': {'en': 'Leggings', 'de': 'Leggings', 'es': 'Leggins', 'nl': 'Leggings', 'sv': 'Leggings', 'fr': 'Leggings', 'da': 'Leggings'},
 'opt2': {'en': 'Tights', 'de': 'Strumpfhose', 'es': 'Medias', 'nl': 'Panty', 'sv': 'Strumpbyxor', 'fr': 'Collants', 'da': 'Strømpebukser'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic cartoon features a coyote chasing a fast-running bird?', 'de': 'In welchem klassischen Cartoon jagt ein Kojote einen schnell laufenden Vogel?', 'es': '¿Que dibujo animado clasico presenta a un coyote persiguiendo a un ave veloz?', 'nl': 'In welke klassieke tekenfilm jaagt een coyote op een snel rennende vogel?', 'sv': 'I vilken klassisk tecknad film jagar en prarievarg en snabbspringande fagel?', 'fr': "Quel dessin anime classique met en scene un coyote poursuivant un oiseau rapide ?", 'da': 'Hvilken klassisk tegnefilm har en coyote, der jagter en hurtigløbende fugl?'},
 'opt1': {'en': 'Wile E. Coyote and the Road Runner', 'de': 'Wile E. Kojote und der Roadrunner', 'es': 'Coyote y Correcaminos', 'nl': 'Wile E. Coyote en de Road Runner', 'sv': 'Coyoten och Vagsprinten', 'fr': 'Vil Coyote et Bip Bip', 'da': 'Wile E. Coyote og Road Runner'},
 'opt2': {'en': 'Tom and Jerry', 'de': 'Tom und Jerry', 'es': 'Tom y Jerry', 'nl': 'Tom en Jerry', 'sv': 'Tom och Jerry', 'fr': 'Tom et Jerry', 'da': 'Tom og Jerry'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the Great Rift Valley?', 'de': 'In welchem Land liegt der Grosse Ostafrikanische Grabenbruch?', 'es': '¿En que pais esta el Gran Valle del Rift?', 'nl': 'In welk land ligt de Grote Slenk?', 'sv': 'I vilket land ligger Stora Riftdalen?', 'fr': "Dans quel pays se trouve le Grand Rift ?", 'da': 'Hvilket land huser Den Store Rift Dal?'},
 'opt1': {'en': 'Kenya', 'de': 'Kenia', 'es': 'Kenia', 'nl': 'Kenia', 'sv': 'Kenya', 'fr': 'le Kenya', 'da': 'Kenya'},
 'opt2': {'en': 'Nigeria', 'de': 'Nigeria', 'es': 'Nigeria', 'nl': 'Nigeria', 'sv': 'Nigeria', 'fr': 'le Nigeria', 'da': 'Nigeria'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which First Lady was known for her 'Just Say No' anti-drug campaign in the 1980s?", 'de': "Welche First Lady war fuer ihre Anti-Drogen-Kampagne 'Just Say No' in den 1980er Jahren bekannt?", 'es': "¿Que primera dama fue conocida por su campaña antidrogas 'Just Say No' en los años 80?", 'nl': "Welke First Lady stond bekend om haar antidrugscampagne 'Just Say No' in de jaren 80?", 'sv': "Vilken First Lady var kand for sin antidrogkampanj 'Just Say No' pa 1980-talet?", 'fr': "Quelle Première dame etait connue pour sa campagne anti-drogue 'Just Say No' dans les annees 1980 ?", 'da': "Hvilken førstedame var kendt for sin anti-narkotika-kampagne 'Just Say No' i 1980’erne?"},
 'opt1': {'en': 'Nancy Reagan', 'de': 'Nancy Reagan', 'es': 'Nancy Reagan', 'nl': 'Nancy Reagan', 'sv': 'Nancy Reagan', 'fr': 'Nancy Reagan', 'da': 'Nancy Reagan'},
 'opt2': {'en': 'Rosalynn Carter', 'de': 'Rosalynn Carter', 'es': 'Rosalynn Carter', 'nl': 'Rosalynn Carter', 'sv': 'Rosalynn Carter', 'fr': 'Rosalynn Carter', 'da': 'Rosalynn Carter'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Brazilian dish is a black bean stew traditionally served with rice?', 'de': 'Welches brasilianische Gericht ist ein Eintopf aus schwarzen Bohnen, traditionell mit Reis serviert?', 'es': '¿Que plato brasileño es un guiso de frijoles negros tradicionalmente servido con arroz?', 'nl': 'Welk Braziliaans gerecht is een stoofpot van zwarte bonen, traditioneel geserveerd met rijst?', 'sv': 'Vilken brasiliansk ratt ar en gryta med svarta bonor, traditionellt serverad med ris?', 'fr': "Quel plat bresilien est un ragout de haricots noirs traditionnellement servi avec du riz ?", 'da': 'Hvilken brasiliansk ret er en gryde med sorte bønner, traditionelt serveret med ris?'},
 'opt1': {'en': 'Feijoada', 'de': 'Feijoada', 'es': 'Feijoada', 'nl': 'Feijoada', 'sv': 'Feijoada', 'fr': 'Feijoada', 'da': 'Feijoada'},
 'opt2': {'en': 'Moqueca', 'de': 'Moqueca', 'es': 'Moqueca', 'nl': 'Moqueca', 'sv': 'Moqueca', 'fr': 'Moqueca', 'da': 'Moqueca'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which part of the ear helps maintain balance?', 'de': 'Welcher Teil des Ohrs hilft, das Gleichgewicht zu halten?', 'es': '¿Que parte del oido ayuda a mantener el equilibrio?', 'nl': 'Welk deel van het oor helpt het evenwicht te behouden?', 'sv': 'Vilken del av orat hjalper till att behalla balansen?', 'fr': "Quelle partie de l'oreille aide a maintenir l'equilibre ?", 'da': 'Hvilken del af øret hjælper med at bevare balancen?'},
 'opt1': {'en': 'Eardrum', 'de': 'Trommelfell', 'es': 'Timpano', 'nl': 'Trommelvlies', 'sv': 'Trumhinnan', 'fr': 'Le tympan', 'da': 'Trommehinden'},
 'opt2': {'en': 'Inner ear (vestibular system)', 'de': 'Innenohr (Vestibularsystem)', 'es': 'Oido interno (sistema vestibular)', 'nl': 'Binnenoor (vestibulair systeem)', 'sv': 'Innerorat (vestibulara systemet)', 'fr': "Oreille interne (systeme vestibulaire)", 'da': 'Det indre øre (vestibulærsystemet)'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which author wrote 'Beloved'?", 'de': "Welche Autorin schrieb 'Menschenkind'?", 'es': "¿Que autora escribio 'Amada'?", 'nl': "Welke auteur schreef 'Beloved'?", 'sv': "Vilken forfattare skrev 'Alskade'?", 'fr': "Quelle autrice a ecrit 'Beloved' ?", 'da': "Hvilken forfatter skrev 'Elskede'?"},
 'opt1': {'en': 'Toni Morrison', 'de': 'Toni Morrison', 'es': 'Toni Morrison', 'nl': 'Toni Morrison', 'sv': 'Toni Morrison', 'fr': 'Toni Morrison', 'da': 'Toni Morrison'},
 'opt2': {'en': 'Alice Walker', 'de': 'Alice Walker', 'es': 'Alice Walker', 'nl': 'Alice Walker', 'sv': 'Alice Walker', 'fr': 'Alice Walker', 'da': 'Alice Walker'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Fantine in the 2012 film adaptation of 'Les Miserables'?", 'de': "Welche Schauspielerin spielte Fantine in der Filmadaption von 'Les Miserables' 2012?", 'es': "¿Que actriz interpreto a Fantine en la adaptacion cinematografica de 2012 de 'Los Miserables'?", 'nl': "Welke actrice speelde Fantine in de filmadaptatie van 'Les Miserables' uit 2012?", 'sv': "Vilken skadespelerska spelade Fantine i 2012 ars filmatisering av 'Les Miserables'?", 'fr': "Quelle actrice a joue Fantine dans l'adaptation cinematographique de 2012 des 'Miserables' ?", 'da': "Hvilken skuespiller spillede Fantine i 2012-filmatiseringen af 'Les Miserables'?"},
 'opt1': {'en': 'Amanda Seyfried', 'de': 'Amanda Seyfried', 'es': 'Amanda Seyfried', 'nl': 'Amanda Seyfried', 'sv': 'Amanda Seyfried', 'fr': 'Amanda Seyfried', 'da': 'Amanda Seyfried'},
 'opt2': {'en': 'Anne Hathaway', 'de': 'Anne Hathaway', 'es': 'Anne Hathaway', 'nl': 'Anne Hathaway', 'sv': 'Anne Hathaway', 'fr': 'Anne Hathaway', 'da': 'Anne Hathaway'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which singer released the album 'Tapestry' in 1971, a landmark of singer-songwriter music?", 'de': "Welche Saengerin veroeffentlichte 1971 das Album 'Tapestry', ein Meilenstein der Singer-Songwriter-Musik?", 'es': "¿Que cantante lanzo el album 'Tapestry' en 1971, un hito de la musica de cantautores?", 'nl': "Welke zangeres bracht in 1971 het album 'Tapestry' uit, een mijlpaal van singer-songwritermuziek?", 'sv': "Vilken sangerska slappte albumet 'Tapestry' 1971, en milstolpe inom singer-songwriter-musik?", 'fr': "Quelle chanteuse a sorti l'album 'Tapestry' en 1971, un jalon de la musique auteur-compositeur-interprete ?", 'da': "Hvilken sanger udgav albummet 'Tapestry' i 1971, en milepæl inden for singer-songwriter-musik?"},
 'opt1': {'en': 'Joni Mitchell', 'de': 'Joni Mitchell', 'es': 'Joni Mitchell', 'nl': 'Joni Mitchell', 'sv': 'Joni Mitchell', 'fr': 'Joni Mitchell', 'da': 'Joni Mitchell'},
 'opt2': {'en': 'Carole King', 'de': 'Carole King', 'es': 'Carole King', 'nl': 'Carole King', 'sv': 'Carole King', 'fr': 'Carole King', 'da': 'Carole King'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Timbuktu?', 'de': 'In welchem Land liegt die antike Stadt Timbuktu?', 'es': '¿En que pais esta la antigua ciudad de Tombuctu?', 'nl': 'In welk land ligt de oude stad Timboektoe?', 'sv': 'I vilket land ligger den antika staden Timbuktu?', 'fr': "Dans quel pays se trouve l'ancienne cite de Tombouctou ?", 'da': 'Hvilket land huser den antikke by Timbuktu?'},
 'opt1': {'en': 'Mali', 'de': 'Mali', 'es': 'Mali', 'nl': 'Mali', 'sv': 'Mali', 'fr': 'le Mali', 'da': 'Mali'},
 'opt2': {'en': 'Niger', 'de': 'Niger', 'es': 'Niger', 'nl': 'Niger', 'sv': 'Niger', 'fr': 'le Niger', 'da': 'Niger'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large Arctic bear is the largest land carnivore?', 'de': 'Welcher grosse arktische Baer ist das groesste Landraubtier?', 'es': '¿Que gran oso artico es el mayor carnivoro terrestre?', 'nl': 'Welke grote arctische beer is het grootste landroofdier?', 'sv': 'Vilken stor arktisk bjorn ar det storsta rovdjuret pa land?', 'fr': "Quel grand ours arctique est le plus grand carnivore terrestre ?", 'da': 'Hvilken stor arktisk bjørn er det største rovdyr på land?'},
 'opt1': {'en': 'Polar bear', 'de': 'Eisbaer', 'es': 'Oso polar', 'nl': 'IJsbeer', 'sv': 'Isbjorn', 'fr': 'Ours polaire', 'da': 'Isbjørn'},
 'opt2': {'en': 'Grizzly bear', 'de': 'Grizzlybaer', 'es': 'Oso grizzly', 'nl': 'Grizzlybeer', 'sv': 'Grizzlybjorn', 'fr': 'Ours grizzli', 'da': 'Grizzlybjørn'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Turkish dessert consists of layers of phyllo dough, nuts, and honey?', 'de': 'Welches tuerkische Dessert besteht aus Schichten von Phyllo-Teig, Nuessen und Honig?', 'es': '¿Que postre turco consiste en capas de masa filo, nueces y miel?', 'nl': 'Welk Turks dessert bestaat uit lagen filodeeg, noten en honing?', 'sv': 'Vilken turkisk dessert bestar av lager av fyllodeg, notter och honung?', 'fr': "Quel dessert turc est compose de couches de pate phyllo, de noix et de miel ?", 'da': 'Hvilken tyrkisk dessert består af lag af filodej, nødder og honning?'},
 'opt1': {'en': 'Baklava', 'de': 'Baklava', 'es': 'Baklava', 'nl': 'Baklava', 'sv': 'Baklava', 'fr': 'Baklava', 'da': 'Baklava'},
 'opt2': {'en': 'Kunefe', 'de': 'Kuenefe', 'es': 'Kunefe', 'nl': 'Kunefe', 'sv': 'Kunefe', 'fr': 'Kunefe', 'da': 'Kunefe'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was a former U.S. Senator and Secretary of State?', 'de': 'Welche First Lady war eine ehemalige US-Senatorin und Außenministerin?', 'es': '¿Que primera dama fue una ex senadora de EE.UU. y secretaria de Estado?', 'nl': 'Welke First Lady was een voormalige Amerikaanse senator en minister van Buitenlandse Zaken?', 'sv': 'Vilken First Lady var en tidigare amerikansk senator och utrikesminister?', 'fr': "Quelle Première dame a ete ancienne senatrice americaine et secretaire d'État ?", 'da': 'Hvilken førstedame var en tidligere amerikansk senator og udenrigsminister?'},
 'opt1': {'en': 'Hillary Clinton', 'de': 'Hillary Clinton', 'es': 'Hillary Clinton', 'nl': 'Hillary Clinton', 'sv': 'Hillary Clinton', 'fr': 'Hillary Clinton', 'da': 'Hillary Clinton'},
 'opt2': {'en': 'Elizabeth Dole', 'de': 'Elizabeth Dole', 'es': 'Elizabeth Dole', 'nl': 'Elizabeth Dole', 'sv': 'Elizabeth Dole', 'fr': 'Elizabeth Dole', 'da': 'Elizabeth Dole'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which artist created the 'Blue Period' and 'Rose Period' before Cubism?", 'de': "Welcher Kuenstler schuf die 'Blaue Periode' und die 'Rosa Periode' vor dem Kubismus?", 'es': "¿Que artista creo el 'Periodo Azul' y el 'Periodo Rosa' antes del cubismo?", 'nl': "Welke kunstenaar creëerde de 'Blauwe Periode' en 'Roze Periode' voor het kubisme?", 'sv': "Vilken konstnar skapade den 'Blaa perioden' och den 'Rosa perioden' fore kubismen?", 'fr': "Quel artiste a cree la 'Periode bleue' et la 'Periode rose' avant le cubisme ?", 'da': 'Hvilken kunstner skabte den "blå periode" og "lyserøde periode" før kubismen?'},
 'opt1': {'en': 'Pablo Picasso', 'de': 'Pablo Picasso', 'es': 'Pablo Picasso', 'nl': 'Pablo Picasso', 'sv': 'Pablo Picasso', 'fr': 'Pablo Picasso', 'da': 'Pablo Picasso'},
 'opt2': {'en': 'Henri Matisse', 'de': 'Henri Matisse', 'es': 'Henri Matisse', 'nl': 'Henri Matisse', 'sv': 'Henri Matisse', 'fr': 'Henri Matisse', 'da': 'Henri Matisse'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which flowering plant is known for its trumpet-shaped blooms and is toxic if ingested?', 'de': 'Welche bluehende Pflanze ist fuer ihre trompetenfoermigen Blueten bekannt und giftig, wenn sie verzehrt wird?', 'es': '¿Que planta con flores es conocida por sus flores en forma de trompeta y es toxica si se ingiere?', 'nl': 'Welke bloeiende plant staat bekend om zijn trompetvormige bloesems en is giftig bij inname?', 'sv': 'Vilken blommande vaxt ar kand for sina trumpetformade blommor och ar giftig om den ats?', 'fr': "Quelle plante a fleurs est connue pour ses fleurs en forme de trompette et est toxique si ingeree ?", 'da': 'Hvilken blomstrende plante er kendt for sine trompetformede blomster og er giftig, hvis den indtages?'},
 'opt1': {'en': 'Lily of the valley', 'de': 'Maiglockchen', 'es': 'Lirio de los valles', 'nl': 'Lelietje-van-dalen', 'sv': 'Liljekonvalj', 'fr': 'Muguet', 'da': 'Liljekonval'},
 'opt2': {'en': 'Foxglove', 'de': 'Fingerhut', 'es': 'Dedalera', 'nl': 'Vingerhoedskruid', 'sv': 'Fingerborgsblomma', 'fr': 'Digitale', 'da': 'Fingerbøl'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which musical is based on the ABBA songbook?', 'de': 'Welches Musical basiert auf dem ABBA-Songbuch?', 'es': '¿Que musical se basa en el repertorio de ABBA?', 'nl': 'Welke musical is gebaseerd op het ABBA-repertoire?', 'sv': 'Vilken musikal ar baserad pa ABBAs latkatalog?', 'fr': "Quelle comedie musicale s'inspire du repertoire d'ABBA ?", 'da': 'Hvilken musical er baseret på ABBAs sangkatalog?'},
 'opt1': {'en': 'Mamma Mia!', 'de': 'Mamma Mia!', 'es': '¡Mamma Mia!', 'nl': 'Mamma Mia!', 'sv': 'Mamma Mia!', 'fr': 'Mamma Mia!', 'da': 'Mamma Mia!'},
 'opt2': {'en': 'Rock of Ages', 'de': 'Rock of Ages', 'es': 'Rock of Ages', 'nl': 'Rock of Ages', 'sv': 'Rock of Ages', 'fr': 'Rock of Ages', 'da': 'Rock of Ages'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country has Prague as its capital?', 'de': 'Welches Land hat Prag als Hauptstadt?', 'es': '¿Que pais tiene a Praga como su capital?', 'nl': 'Welk land heeft Praag als hoofdstad?', 'sv': 'Vilket land har Prag som huvudstad?', 'fr': "Quel pays a Prague pour capitale ?", 'da': 'Hvilket land har Prag som hovedstad?'},
 'opt1': {'en': 'Czech Republic', 'de': 'Tschechien', 'es': 'Republica Checa', 'nl': 'Tsjechie', 'sv': 'Tjeckien', 'fr': 'la Republique tcheque', 'da': 'Tjekkiet'},
 'opt2': {'en': 'Slovakia', 'de': 'Slowakei', 'es': 'Eslovaquia', 'nl': 'Slowakije', 'sv': 'Slovakien', 'fr': 'la Slovaquie', 'da': 'Slovakiet'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which spice is used to flavor traditional root beer and comes from a tree's bark?", 'de': 'Welches Gewuerz wird verwendet, um traditionelles Root Beer zu aromatisieren, und stammt von Baumrinde?', 'es': '¿Que especia se usa para saborizar la root beer tradicional y proviene de la corteza de un arbol?', 'nl': 'Welke specerij wordt gebruikt om traditionele root beer te smaken en komt van boombast?', 'sv': 'Vilken krydda anvands for att smaksatta traditionell root beer och kommer fran trabark?', 'fr': "Quelle epice est utilisee pour aromatiser la root beer traditionnelle et provient de l'ecorce d'un arbre ?", 'da': 'Hvilket krydderi bruges til at smage traditionel root beer og kommer fra et træs bark?'},
 'opt1': {'en': 'Sassafras', 'de': 'Sassafras', 'es': 'Sasafras', 'nl': 'Sassafras', 'sv': 'Sassafras', 'fr': 'Sassafras', 'da': 'Sassafras'},
 'opt2': {'en': 'Cinnamon', 'de': 'Zimt', 'es': 'Canela', 'nl': 'Kaneel', 'sv': 'Kanel', 'fr': 'La cannelle', 'da': 'Kanel'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which craft involves dyeing fabric using wax-resist techniques, popular in Indonesia?', 'de': 'Bei welchem Handwerk wird Stoff mit Wachsreservetechniken gefaerbt, beliebt in Indonesien?', 'es': '¿Que manualidad implica teñir tela usando tecnicas de reserva con cera, popular en Indonesia?', 'nl': 'Bij welk handwerk wordt stof geverfd met wasreservetechnieken, populair in Indonesie?', 'sv': 'Vilket hantverk innebar att farga tyg med vaxreservteknik, populart i Indonesien?', 'fr': "Quel art manuel consiste a teindre du tissu avec des techniques de reserve a la cire, populaire en Indonesie ?", 'da': 'Hvilket håndværk involverer at farve stof med voks-reserveteknikker, populært i Indonesien?'},
 'opt1': {'en': 'Tie-dye', 'de': 'Batikfärben mit Bindetechnik', 'es': 'Teñido anudado', 'nl': 'Tie-dye', 'sv': 'Knytfargning', 'fr': 'Tie-dye', 'da': 'Tie-dye'},
 'opt2': {'en': 'Batik', 'de': 'Batik', 'es': 'Batik', 'nl': 'Batik', 'sv': 'Batik', 'fr': 'Le batik', 'da': 'Batik'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which organ filters toxins from the blood and produces bile?', 'de': 'Welches Organ filtert Toxine aus dem Blut und produziert Galle?', 'es': '¿Que organo filtra toxinas de la sangre y produce bilis?', 'nl': 'Welk orgaan filtert gifstoffen uit het bloed en produceert gal?', 'sv': 'Vilket organ filtrerar gifter fran blodet och producerar galla?', 'fr': "Quel organe filtre les toxines du sang et produit la bile ?", 'da': 'Hvilket organ filtrerer giftstoffer fra blodet og producerer galde?'},
 'opt1': {'en': 'Liver', 'de': 'Leber', 'es': 'Higado', 'nl': 'Lever', 'sv': 'Lever', 'fr': 'Foie', 'da': 'Lever'},
 'opt2': {'en': 'Kidney', 'de': 'Niere', 'es': 'Riñon', 'nl': 'Nier', 'sv': 'Njure', 'fr': 'Rein', 'da': 'Nyre'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'In volleyball, what is the term for a shot that lands just over the net, hard to return?', 'de': 'Wie nennt man beim Volleyball einen Schlag, der knapp ueber das Netz geht und schwer zurueckzuschlagen ist?', 'es': '¿En voleibol, como se llama un golpe que cae justo sobre la red, dificil de devolver?', 'nl': 'Hoe heet bij volleybal een slag die net over het net belandt en moeilijk terug te slaan is?', 'sv': 'Vad kallas i volleyboll ett slag som landar precis over natet, svart att returnera?', 'fr': "En volley-ball, comment appelle-t-on un coup qui atterrit juste au-dessus du filet, difficile a renvoyer ?", 'da': 'Hvad kaldes i volleyball et slag, der lander lige over nettet, svært at returnere?'},
 'opt1': {'en': 'Spike', 'de': 'Angriffsschlag', 'es': 'Remate', 'nl': 'Smash', 'sv': 'Smash', 'fr': 'Smash', 'da': 'Smash'},
 'opt2': {'en': 'Dig', 'de': 'Abwehr', 'es': 'Defensa', 'nl': 'Verdedigende pass', 'sv': 'Grävning', 'fr': 'La manchette', 'da': 'Forsvarsslag'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Italian region is famous for Tuscan wine and rolling hills?', 'de': 'Welche italienische Region ist beruehmt fuer toskanischen Wein und sanfte Huegel?', 'es': '¿Que region italiana es famosa por el vino toscano y las colinas onduladas?', 'nl': 'Welke Italiaanse regio staat bekend om Toscaanse wijn en glooiende heuvels?', 'sv': 'Vilken italiensk region ar kand for toskansk vin och böljande kullar?', 'fr': "Quelle region italienne est celebre pour le vin toscan et ses collines vallonnees ?", 'da': 'Hvilken italiensk region er berømt for toskansk vin og bølgende bakker?'},
 'opt1': {'en': 'Tuscany', 'de': 'Toskana', 'es': 'Toscana', 'nl': 'Toscane', 'sv': 'Toscana', 'fr': 'la Toscane', 'da': 'Toscana'},
 'opt2': {'en': 'Umbria', 'de': 'Umbrien', 'es': 'Umbria', 'nl': 'Umbrie', 'sv': 'Umbrien', 'fr': "l'Ombrie", 'da': 'Umbrien'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Celie in the 1985 film 'The Color Purple'?", 'de': "Welche Schauspielerin spielte Celie im Film 'Die Farbe Lila' von 1985?", 'es': "¿Que actriz interpreto a Celie en la pelicula de 1985 'El color purpura'?", 'nl': "Welke actrice speelde Celie in de film 'The Color Purple' uit 1985?", 'sv': "Vilken skadespelerska spelade Celie i 1985 ars film 'Fargen lila'?", 'fr': "Quelle actrice a joue Celie dans le film 'La Couleur pourpre' de 1985 ?", 'da': "Hvilken skuespiller spillede Celie i 1985-filmen 'Purpurfarven'?"},
 'opt1': {'en': 'Oprah Winfrey', 'de': 'Oprah Winfrey', 'es': 'Oprah Winfrey', 'nl': 'Oprah Winfrey', 'sv': 'Oprah Winfrey', 'fr': 'Oprah Winfrey', 'da': 'Oprah Winfrey'},
 'opt2': {'en': 'Whoopi Goldberg', 'de': 'Whoopi Goldberg', 'es': 'Whoopi Goldberg', 'nl': 'Whoopi Goldberg', 'sv': 'Whoopi Goldberg', 'fr': 'Whoopi Goldberg', 'da': 'Whoopi Goldberg'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which classic novel features the March sisters, including Jo and Beth?', 'de': 'In welchem Klassiker kommen die March-Schwestern vor, darunter Jo und Beth?', 'es': '¿Que novela clasica presenta a las hermanas March, incluyendo a Jo y Beth?', 'nl': 'In welke klassieke roman komen de zusjes March voor, waaronder Jo en Beth?', 'sv': 'I vilken klassisk roman forekommer systrarna March, inklusive Jo och Beth?', 'fr': "Quel roman classique met en scene les sœurs March, dont Jo et Beth ?", 'da': 'Hvilken klassisk roman har March-søstrene, inklusive Jo og Beth?'},
 'opt1': {'en': 'Little Men', 'de': 'Kleine Maenner', 'es': 'Hombrecitos', 'nl': 'Kleine Mannen', 'sv': 'Unga man', 'fr': 'Les Petits Hommes', 'da': 'Små Mænd'},
 'opt2': {'en': 'Little Women', 'de': 'Little Women', 'es': 'Mujercitas', 'nl': 'Onder Moeders Vleugels', 'sv': 'Unga kvinnor', 'fr': 'Les Filles du docteur March', 'da': 'Unge Piger'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which 1970s singer is known for the song 'Bohemian Rhapsody'?", 'de': "Welcher Saenger der 1970er ist fuer das Lied 'Bohemian Rhapsody' bekannt?", 'es': "¿Que cantante de los 70 es conocido por la cancion 'Bohemian Rhapsody'?", 'nl': "Welke zanger uit de jaren 70 staat bekend om het lied 'Bohemian Rhapsody'?", 'sv': "Vilken sangare fran 1970-talet ar kand for latten 'Bohemian Rhapsody'?", 'fr': "Quel chanteur des annees 1970 est connu pour la chanson 'Bohemian Rhapsody' ?", 'da': "Hvilken 1970’er-sanger er kendt for sangen 'Bohemian Rhapsody'?"},
 'opt1': {'en': 'Elton John', 'de': 'Elton John', 'es': 'Elton John', 'nl': 'Elton John', 'sv': 'Elton John', 'fr': 'Elton John', 'da': 'Elton John'},
 'opt2': {'en': 'Freddie Mercury (Queen)', 'de': 'Freddie Mercury (Queen)', 'es': 'Freddie Mercury (Queen)', 'nl': 'Freddie Mercury (Queen)', 'sv': 'Freddie Mercury (Queen)', 'fr': 'Freddie Mercury (Queen)', 'da': 'Freddie Mercury (Queen)'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Southern U.S. drink is made from fermented corn and aged in barrels?', 'de': 'Welches suddstaatliche Getraenk wird aus fermentiertem Mais hergestellt und in Faessern gereift?', 'es': '¿Que bebida sureña de EE.UU. se hace de maiz fermentado y se envejece en barriles?', 'nl': 'Welke zuidelijk Amerikaanse drank wordt gemaakt van gefermenteerde mais en gerijpt in vaten?', 'sv': 'Vilken sydstatlig dryck gors av fermenterad majs och lagras i fat?', 'fr': "Quelle boisson du Sud des États-Unis est faite de mais fermente et vieillie en fut ?", 'da': 'Hvilken sydstatsdrik er lavet af fermenteret majs og lagret i tønder?'},
 'opt1': {'en': 'Moonshine', 'de': 'Moonshine', 'es': 'Moonshine', 'nl': 'Moonshine', 'sv': 'Moonshine', 'fr': 'Alcool de contrebande', 'da': 'Moonshine'},
 'opt2': {'en': 'Bourbon', 'de': 'Bourbon', 'es': 'Bourbon', 'nl': 'Bourbon', 'sv': 'Bourbon', 'fr': 'Bourbon', 'da': 'Bourbon'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which queen ruled France as regent and was known for the St. Bartholomew's Day events?", 'de': "Welche Koenigin regierte Frankreich als Regentin und war bekannt fuer die Ereignisse der Bartholomaeusnacht?", 'es': "¿Que reina goberno Francia como regente y fue conocida por los sucesos de la matanza de San Bartolome?", 'nl': "Welke koningin regeerde Frankrijk als regentes en was bekend om de gebeurtenissen van de Bartholomeusnacht?", 'sv': "Vilken drottning styrde Frankrike som regent och var kand for Bartolomeinatten?", 'fr': "Quelle reine a regne sur la France en tant que regente et etait connue pour les evenements de la Saint-Barthelemy ?", 'da': 'Hvilken dronning regerede Frankrig som regent og var kendt for Bartholomæusnatten?'},
 'opt1': {'en': "Marie de' Medici", 'de': 'Maria von Medici', 'es': 'Maria de Medici', 'nl': "Maria de' Medici", 'sv': 'Maria av Medici', 'fr': 'Marie de Medicis', 'da': 'Maria af Medici'},
 'opt2': {'en': "Catherine de' Medici", 'de': 'Katharina von Medici', 'es': 'Catalina de Medici', 'nl': "Catharina de' Medici", 'sv': "Katarina av Medici", 'fr': 'Catherine de Medicis', 'da': 'Katarina af Medici'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which spotted big cat is the fastest land animal?', 'de': 'Welche gefleckte Grosskatze ist das schnellste Landtier?', 'es': '¿Que felino manchado es el animal terrestre mas rapido?', 'nl': 'Welke gevlekte grote kat is het snelste landdier?', 'sv': 'Vilken flackig storkatt ar det snabbaste landdjuret?', 'fr': "Quel grand felin tachete est l'animal terrestre le plus rapide ?", 'da': 'Hvilken plettet storkat er det hurtigste landdyr?'},
 'opt1': {'en': 'Cheetah', 'de': 'Gepard', 'es': 'Guepardo', 'nl': 'Jachtluipaard', 'sv': 'Gepard', 'fr': 'Guepard', 'da': 'Gepard'},
 'opt2': {'en': 'Leopard', 'de': 'Leopard', 'es': 'Leopardo', 'nl': 'Luipaard', 'sv': 'Leopard', 'fr': 'Leopard', 'da': 'Leopard'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Cusco, former Inca capital?', 'de': 'In welchem Land liegt die antike Stadt Cusco, die ehemalige Inka-Hauptstadt?', 'es': '¿En que pais esta la antigua ciudad de Cusco, antigua capital inca?', 'nl': 'In welk land ligt de oude stad Cusco, voormalige Inca-hoofdstad?', 'sv': 'I vilket land ligger den antika staden Cusco, forna inkahuvudstaden?', 'fr': "Dans quel pays se trouve l'ancienne cite de Cusco, ancienne capitale inca ?", 'da': 'Hvilket land huser den antikke by Cusco, tidligere inka-hovedstad?'},
 'opt1': {'en': 'Ecuador', 'de': 'Ecuador', 'es': 'Ecuador', 'nl': 'Ecuador', 'sv': 'Ecuador', 'fr': "l'Equateur", 'da': 'Ecuador'},
 'opt2': {'en': 'Peru', 'de': 'Peru', 'es': 'Peru', 'nl': 'Peru', 'sv': 'Peru', 'fr': 'le Perou', 'da': 'Peru'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which type of bag, popularized by Chanel, is worn crossbody with a chain strap?', 'de': 'Welche Art von Tasche, populaer gemacht von Chanel, wird mit Kettenriemen ueber der Schulter getragen?', 'es': '¿Que tipo de bolso, popularizado por Chanel, se lleva cruzado con una correa de cadena?', 'nl': 'Welk type tas, populair gemaakt door Chanel, wordt cross-body gedragen met een kettinghengsel?', 'sv': 'Vilken typ av vaska, popularaiserad av Chanel, baers over kroppen med en kedjerem?', 'fr': "Quel type de sac, popularise par Chanel, se porte en bandouliere avec une chaine ?", 'da': 'Hvilken type taske, populariseret af Chanel, bæres over skulderen med en kædestrop?'},
 'opt1': {'en': 'Kelly bag', 'de': 'Kelly Bag', 'es': 'Bolso Kelly', 'nl': 'Kelly bag', 'sv': 'Kelly bag', 'fr': 'Sac Kelly', 'da': 'Kelly-taske'},
 'opt2': {'en': 'Chanel 2.55', 'de': 'Chanel 2.55', 'es': 'Chanel 2.55', 'nl': 'Chanel 2.55', 'sv': 'Chanel 2.55', 'fr': 'Chanel 2.55', 'da': 'Chanel 2.55'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which teen drama series follows the wealthy Waldorf and van der Woodsen social circle at a Manhattan prep school?', 'de': 'Welche Teenie-Drama-Serie folgt dem wohlhabenden sozialen Kreis von Waldorf und van der Woodsen an einer Manhattan-Privatschule?', 'es': '¿Que serie de drama adolescente sigue al circulo social adinerado de Waldorf y van der Woodsen en una escuela preparatoria de Manhattan?', 'nl': 'Welke tienerdramaserie volgt de rijke sociale kring van Waldorf en van der Woodsen op een Manhattan prepschool?', 'sv': 'Vilken tonarsdramaserie foljer den formogna sociala kretsen Waldorf och van der Woodsen pa en Manhattan-privatskola?', 'fr': "Quelle serie dramatique pour adolescents suit le cercle social fortune de Waldorf et van der Woodsen dans une ecole privee de Manhattan ?", 'da': 'Hvilken teenagedrama-serie følger den velhavende Waldorf og van der Woodsen sociale kreds på en Manhattan privatskole?'},
 'opt1': {'en': 'Emily in Paris', 'de': 'Emily in Paris', 'es': 'Emily en Paris', 'nl': 'Emily in Paris', 'sv': 'Emily i Paris', 'fr': 'Emily in Paris', 'da': 'Emily in Paris'},
 'opt2': {'en': 'Gossip Girl', 'de': 'Gossip Girl', 'es': 'Gossip Girl', 'nl': 'Gossip Girl', 'sv': 'Gossip Girl', 'fr': 'Gossip Girl', 'da': 'Gossip Girl'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which flower is known for its heart-shaped, dangling blooms and is a symbol of love?', 'de': 'Welche Blume ist fuer ihre herzfoermigen, haengenden Blueten bekannt und ist ein Symbol der Liebe?', 'es': '¿Que flor es conocida por sus flores colgantes en forma de corazon y es simbolo del amor?', 'nl': 'Welke bloem staat bekend om zijn hartvormige, hangende bloesems en is een liefdessymbool?', 'sv': 'Vilken blomma ar kand for sina hjartformade, hangande blommor och ar en karlekssymbol?', 'fr': "Quelle fleur est connue pour ses fleurs pendantes en forme de coeur et est un symbole d'amour ?", 'da': 'Hvilken blomst er kendt for sine hjerteformede, hængende blomster og er et kærlighedssymbol?'},
 'opt1': {'en': 'Bleeding heart', 'de': 'Traenendes Herz', 'es': 'Corazon sangrante', 'nl': 'Gebroken hartje', 'sv': 'Löjtnantshjarta', 'fr': "Coeur de Marie", 'da': 'Løjtnantshjerte'},
 'opt2': {'en': 'Forget-me-not', 'de': 'Vergissmeinnicht', 'es': 'Nomeolvides', 'nl': 'Vergeet-mij-nietje', 'sv': 'Fortgatmigej', 'fr': "Myosotis", 'da': 'Forglemmigej'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Filipino dish is chicken or pork marinated in vinegar, soy sauce, and garlic?', 'de': 'Welches philippinische Gericht ist Huhn oder Schwein, mariniert in Essig, Sojasauce und Knoblauch?', 'es': '¿Que plato filipino es pollo o cerdo marinado en vinagre, salsa de soja y ajo?', 'nl': 'Welk Filipijns gerecht is kip of varkensvlees gemarineerd in azijn, sojasaus en knoflook?', 'sv': 'Vilken filippinsk ratt ar kyckling eller flask marinerad i vinager, sojasas och vitlok?', 'fr': "Quel plat philippin est du poulet ou du porc marine dans du vinaigre, de la sauce soja et de l'ail ?", 'da': 'Hvilken filippinsk ret er kylling eller svinekød marineret i eddike, sojasovs og hvidløg?'},
 'opt1': {'en': 'Sinigang', 'de': 'Sinigang', 'es': 'Sinigang', 'nl': 'Sinigang', 'sv': 'Sinigang', 'fr': 'Sinigang', 'da': 'Sinigang'},
 'opt2': {'en': 'Adobo', 'de': 'Adobo', 'es': 'Adobo', 'nl': 'Adobo', 'sv': 'Adobo', 'fr': 'Adobo', 'da': 'Adobo'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which body part regulates body temperature through sweating?', 'de': 'Welcher Koerperteil reguliert die Koerpertemperatur durch Schwitzen?', 'es': '¿Que parte del cuerpo regula la temperatura corporal mediante la sudoracion?', 'nl': 'Welk lichaamsdeel reguleert de lichaamstemperatuur door zweten?', 'sv': 'Vilken kroppsdel reglerar kroppstemperaturen genom svettning?', 'fr': "Quelle partie du corps regule la temperature corporelle par la transpiration ?", 'da': 'Hvilken kropsdel regulerer kropstemperaturen gennem svedning?'},
 'opt1': {'en': 'Skin', 'de': 'Haut', 'es': 'Piel', 'nl': 'Huid', 'sv': 'Hud', 'fr': 'La peau', 'da': 'Hud'},
 'opt2': {'en': 'Liver', 'de': 'Leber', 'es': 'Higado', 'nl': 'Lever', 'sv': 'Lever', 'fr': 'Foie', 'da': 'Lever'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Luxor and the Valley of the Kings?', 'de': 'In welchem Land liegen die antike Stadt Luxor und das Tal der Koenige?', 'es': '¿En que pais estan la antigua ciudad de Luxor y el Valle de los Reyes?', 'nl': 'In welk land liggen de oude stad Luxor en de Vallei der Koningen?', 'sv': 'I vilket land ligger den antika staden Luxor och Konungarnas dal?', 'fr': "Dans quel pays se trouvent l'ancienne cite de Louxor et la Vallee des Rois ?", 'da': 'Hvilket land huser den antikke by Luxor og Kongernes Dal?'},
 'opt1': {'en': 'Egypt', 'de': 'Ägypten', 'es': 'Egipto', 'nl': 'Egypte', 'sv': 'Egypten', 'fr': "l'Egypte", 'da': 'Egypten'},
 'opt2': {'en': 'Sudan', 'de': 'Sudan', 'es': 'Sudan', 'nl': 'Soedan', 'sv': 'Sudan', 'fr': 'le Soudan', 'da': 'Sudan'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was known for her needlepoint and support of the Foster Grandparent Program?', 'de': 'Welche First Lady war bekannt fuer ihre Handarbeit und Unterstuetzung des Foster Grandparent Program?', 'es': '¿Que primera dama fue conocida por su bordado de aguja y apoyo al Programa de Abuelos Adoptivos?', 'nl': 'Welke First Lady stond bekend om haar borduurwerk en steun aan het Foster Grandparent Program?', 'sv': 'Vilken First Lady var kand for sitt nalarbete och stod for Foster Grandparent Program?', 'fr': "Quelle Première dame etait connue pour son travail a l'aiguille et son soutien au Foster Grandparent Program ?", 'da': 'Hvilken førstedame var kendt for sit nålearbejde og støtte til Foster Grandparent Program?'},
 'opt1': {'en': 'Rosalynn Carter', 'de': 'Rosalynn Carter', 'es': 'Rosalynn Carter', 'nl': 'Rosalynn Carter', 'sv': 'Rosalynn Carter', 'fr': 'Rosalynn Carter', 'da': 'Rosalynn Carter'},
 'opt2': {'en': 'Betty Ford', 'de': 'Betty Ford', 'es': 'Betty Ford', 'nl': 'Betty Ford', 'sv': 'Betty Ford', 'fr': 'Betty Ford', 'da': 'Betty Ford'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which composer wrote 'Peter and the Wolf', a work for children with narrated instruments?", 'de': "Welcher Komponist schrieb 'Peter und der Wolf', ein Kinderwerk mit erzaehlten Instrumenten?", 'es': "¿Que compositor escribio 'Pedro y el lobo', una obra para niños con instrumentos narrados?", 'nl': "Welke componist schreef 'Peter en de Wolf', een werk voor kinderen met verteld instrumentarium?", 'sv': "Vilken kompositor skrev 'Peter och vargen', ett verk for barn med berattade instrument?", 'fr': "Quel compositeur a ecrit 'Pierre et le Loup', une œuvre pour enfants avec instruments narres ?", 'da': "Hvilken komponist skrev 'Peter og ulven', et værk for børn med fortalte instrumenter?"},
 'opt1': {'en': 'Igor Stravinsky', 'de': 'Igor Strawinsky', 'es': 'Igor Stravinski', 'nl': 'Igor Stravinsky', 'sv': 'Igor Stravinskij', 'fr': 'Igor Stravinski', 'da': 'Igor Stravinskij'},
 'opt2': {'en': 'Sergei Prokofiev', 'de': 'Sergei Prokofjew', 'es': 'Sergei Prokofiev', 'nl': 'Sergej Prokofjev', 'sv': 'Sergej Prokofjev', 'fr': 'Sergueï Prokofiev', 'da': 'Sergej Prokofjev'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Sally Bowles in the film 'Cabaret'?", 'de': "Welche Schauspielerin spielte Sally Bowles im Film 'Cabaret'?", 'es': "¿Que actriz interpreto a Sally Bowles en la pelicula 'Cabaret'?", 'nl': "Welke actrice speelde Sally Bowles in de film 'Cabaret'?", 'sv': "Vilken skadespelerska spelade Sally Bowles i filmen 'Cabaret'?", 'fr': "Quelle actrice a joue Sally Bowles dans le film 'Cabaret' ?", 'da': "Hvilken skuespiller spillede Sally Bowles i filmen 'Cabaret'?"},
 'opt1': {'en': 'Judy Garland', 'de': 'Judy Garland', 'es': 'Judy Garland', 'nl': 'Judy Garland', 'sv': 'Judy Garland', 'fr': 'Judy Garland', 'da': 'Judy Garland'},
 'opt2': {'en': 'Liza Minnelli', 'de': 'Liza Minnelli', 'es': 'Liza Minnelli', 'nl': 'Liza Minnelli', 'sv': 'Liza Minnelli', 'fr': 'Liza Minnelli', 'da': 'Liza Minnelli'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Spanish rice dish is cooked with saffron and often includes seafood?', 'de': 'Welches spanische Reisgericht wird mit Safran gekocht und enthaelt oft Meeresfruechte?', 'es': '¿Que plato de arroz español se cocina con azafran y a menudo incluye mariscos?', 'nl': 'Welk Spaans rijstgerecht wordt gekookt met saffraan en bevat vaak zeevruchten?', 'sv': 'Vilken spansk risratt tillagas med saffran och innehaller ofta skaldjur?', 'fr': "Quel plat de riz espagnol est cuisine au safran et comprend souvent des fruits de mer ?", 'da': 'Hvilken spansk risret tilberedes med safran og indeholder ofte skaldyr?'},
 'opt1': {'en': 'Risotto', 'de': 'Risotto', 'es': 'Risotto', 'nl': 'Risotto', 'sv': 'Risotto', 'fr': 'Risotto', 'da': 'Risotto'},
 'opt2': {'en': 'Paella', 'de': 'Paella', 'es': 'Paella', 'nl': 'Paella', 'sv': 'Paella', 'fr': 'Paella', 'da': 'Paella'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to Niagara Falls (on the American side)?', 'de': 'In welchem US-Bundesstaat liegen die Niagarafaelle (auf der amerikanischen Seite)?', 'es': '¿En que estado de EE.UU. estan las Cataratas del Niagara (en el lado americano)?', 'nl': 'In welke Amerikaanse staat liggen Niagara Falls (aan de Amerikaanse kant)?', 'sv': 'I vilken amerikansk delstat ligger Niagarafallen (pa den amerikanska sidan)?', 'fr': "Dans quel État americain se trouvent les chutes du Niagara (cote americain) ?", 'da': 'I hvilken amerikansk delstat ligger Niagara Falls (på den amerikanske side)?'},
 'opt1': {'en': 'New York', 'de': 'New York', 'es': 'Nueva York', 'nl': 'New York', 'sv': 'New York', 'fr': "l'État de New York", 'da': 'New York'},
 'opt2': {'en': 'Vermont', 'de': 'Vermont', 'es': 'Vermont', 'nl': 'Vermont', 'sv': 'Vermont', 'fr': 'le Vermont', 'da': 'Vermont'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which small mammal is known for its bushy tail and burying acorns?', 'de': 'Welches kleine Saeugetier ist fuer seinen buschigen Schwanz und das Vergraben von Eicheln bekannt?', 'es': '¿Que pequeño mamifero es conocido por su cola tupida y por enterrar bellotas?', 'nl': 'Welk klein zoogdier staat bekend om zijn borstelige staart en het begraven van eikels?', 'sv': 'Vilket litet daggdjur ar kant for sin buskiga svans och att grava ner ekollon?', 'fr': "Quel petit mammifere est connu pour sa queue touffue et le fait d'enterrer des glands ?", 'da': 'Hvilket lille pattedyr er kendt for sin buskede hale og at begrave agern?'},
 'opt1': {'en': 'Chipmunk', 'de': 'Streifenhoernchen', 'es': 'Ardilla listada', 'nl': 'Wangzakeekhoorn', 'sv': 'Jordekorre', 'fr': 'Tamia', 'da': 'Jordegern'},
 'opt2': {'en': 'Squirrel', 'de': 'Eichhoernchen', 'es': 'Ardilla', 'nl': 'Eekhoorn', 'sv': 'Ekorre', 'fr': 'Ecureuil', 'da': 'Egern'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which German Christmas market tradition features mulled wine and handmade crafts?', 'de': 'Welche deutsche Weihnachtsmarkt-Tradition bietet Gluehwein und handgefertigte Handwerkskunst?', 'es': '¿Que tradicion de mercado navideño aleman presenta vino caliente y artesania hecha a mano?', 'nl': 'Welke Duitse kerstmarkttraditie biedt gluhwein en handgemaakte ambachten?', 'sv': 'Vilken tysk julmarknadstradition bjuder pa glogg och handgjorda hantverk?', 'fr': "Quelle tradition de marche de Noel allemand propose du vin chaud et de l'artisanat fait main ?", 'da': 'Hvilken tysk julemarkedstradition byder på gløgg og håndlavet håndværk?'},
 'opt1': {'en': 'Oktoberfest', 'de': 'Oktoberfest', 'es': 'Oktoberfest', 'nl': 'Oktoberfest', 'sv': 'Oktoberfest', 'fr': 'Oktoberfest', 'da': 'Oktoberfest'},
 'opt2': {'en': 'Christkindlmarkt', 'de': 'Christkindlmarkt', 'es': 'Christkindlmarkt', 'nl': 'Christkindlmarkt', 'sv': 'Christkindlmarkt', 'fr': 'Marche de Noel (Christkindlmarkt)', 'da': 'Christkindlmarkt'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'The Joy Luck Club'?", 'de': "Welche Autorin schrieb 'Der Klub der Gluecklichen Frauen'?", 'es': "¿Que autora escribio 'El club de la buena estrella'?", 'nl': "Welke auteur schreef 'The Joy Luck Club'?", 'sv': "Vilken forfattare skrev 'Lyckoklubben'?", 'fr': "Quelle autrice a ecrit 'Le Club de la chance' ?", 'da': "Hvilken forfatter skrev 'The Joy Luck Club'?"},
 'opt1': {'en': 'Maxine Hong Kingston', 'de': 'Maxine Hong Kingston', 'es': 'Maxine Hong Kingston', 'nl': 'Maxine Hong Kingston', 'sv': 'Maxine Hong Kingston', 'fr': 'Maxine Hong Kingston', 'da': 'Maxine Hong Kingston'},
 'opt2': {'en': 'Amy Tan', 'de': 'Amy Tan', 'es': 'Amy Tan', 'nl': 'Amy Tan', 'sv': 'Amy Tan', 'fr': 'Amy Tan', 'da': 'Amy Tan'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Samarkand?', 'de': 'In welchem Land liegt die antike Stadt Samarkand?', 'es': '¿En que pais esta la antigua ciudad de Samarcanda?', 'nl': 'In welk land ligt de oude stad Samarkand?', 'sv': 'I vilket land ligger den antika staden Samarkand?', 'fr': "Dans quel pays se trouve l'ancienne cite de Samarcande ?", 'da': 'Hvilket land huser den antikke by Samarkand?'},
 'opt1': {'en': 'Uzbekistan', 'de': 'Usbekistan', 'es': 'Uzbekistan', 'nl': 'Oezbekistan', 'sv': 'Uzbekistan', 'fr': "l'Ouzbekistan", 'da': 'Usbekistan'},
 'opt2': {'en': 'Turkmenistan', 'de': 'Turkmenistan', 'es': 'Turkmenistan', 'nl': 'Turkmenistan', 'sv': 'Turkmenistan', 'fr': 'le Turkmenistan', 'da': 'Turkmenistan'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which instrument is central to reggae music, providing an offbeat rhythm?', 'de': 'Welches Instrument ist zentral fuer Reggae-Musik und liefert einen Offbeat-Rhythmus?', 'es': '¿Que instrumento es central en la musica reggae, aportando un ritmo a contratiempo?', 'nl': 'Welk instrument is centraal in reggaemuziek en zorgt voor een offbeat-ritme?', 'sv': 'Vilket instrument ar centralt i reggaemusik och ger en offbeat-rytm?', 'fr': "Quel instrument est central dans la musique reggae, fournissant un rythme a contretemps ?", 'da': 'Hvilket instrument er centralt i reggaemusik og giver en offbeat-rytme?'},
 'opt1': {'en': 'Guitar (skank rhythm)', 'de': 'Gitarre (Skank-Rhythmus)', 'es': 'Guitarra (ritmo skank)', 'nl': 'Gitaar (skank-ritme)', 'sv': 'Gitarr (skank-rytm)', 'fr': 'Guitare (rythme skank)', 'da': 'Guitar (skank-rytme)'},
 'opt2': {'en': 'Steel drum', 'de': 'Steel Drum', 'es': 'Tambor de acero', 'nl': 'Steeldrum', 'sv': 'Stalpanna', 'fr': 'Steel drum', 'da': 'Steel drum'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Southern U.S. dish is a rich, cheesy baked pasta casserole?', 'de': 'Welches suddstaatliche Gericht ist ein reichhaltiger, kaesiger Nudelauflauf?', 'es': '¿Que plato sureño de EE.UU. es una cacerola de pasta gratinada rica y con queso?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een rijke, kazige pasta-ovenschotel?', 'sv': 'Vilken sydstatlig ratt ar en rik, ostig ugnsbakad pastagryta?', 'fr': "Quel plat du Sud des États-Unis est un gratin de pates riche et fromager ?", 'da': 'Hvilken sydstatsret er en rig, ostet bagt pastagryde?'},
 'opt1': {'en': 'Lasagna', 'de': 'Lasagne', 'es': 'Lasaña', 'nl': 'Lasagne', 'sv': 'Lasagne', 'fr': 'Lasagne', 'da': 'Lasagne'},
 'opt2': {'en': 'Baked mac and cheese', 'de': 'Baked Mac and Cheese', 'es': 'Macarrones con queso al horno', 'nl': 'Gebakken mac and cheese', 'sv': 'Ugnsbakad mac and cheese', 'fr': 'Macaroni au fromage gratine', 'da': 'Bagt mac and cheese'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which mineral is essential for thyroid function and found in iodized salt?', 'de': 'Welches Mineral ist essentiell fuer die Schilddruesenfunktion und in jodiertem Salz enthalten?', 'es': '¿Que mineral es esencial para la funcion tiroidea y se encuentra en la sal yodada?', 'nl': 'Welk mineraal is essentieel voor de schildklierfunctie en te vinden in gejodeerd zout?', 'sv': 'Vilket mineral ar viktigt for skoldkortelfunktionen och finns i joderat salt?', 'fr': "Quel mineral est essentiel a la fonction thyroidienne et se trouve dans le sel iode ?", 'da': 'Hvilket mineral er essentielt for skjoldbruskkirtelfunktionen og findes i jodberiget salt?'},
 'opt1': {'en': 'Zinc', 'de': 'Zink', 'es': 'Zinc', 'nl': 'Zink', 'sv': 'Zink', 'fr': 'Zinc', 'da': 'Zink'},
 'opt2': {'en': 'Iodine', 'de': 'Jod', 'es': 'Yodo', 'nl': 'Jodium', 'sv': 'Jod', 'fr': 'Iode', 'da': 'Jod'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which film has Julia Roberts playing a lawyer's assistant investigating a water contamination case?", 'de': 'In welchem Film spielt Julia Roberts eine Anwaltsassistentin, die einen Fall von Wasserverschmutzung untersucht?', 'es': '¿En que pelicula Julia Roberts interpreta a una asistente legal que investiga un caso de contaminacion del agua?', 'nl': 'In welke film speelt Julia Roberts een advocatenassistent die een waterverontreinigingszaak onderzoekt?', 'sv': 'I vilken film spelar Julia Roberts en advokatassistent som utreder ett vattenfororeningsfall?', 'fr': "Dans quel film Julia Roberts joue-t-elle une assistante juridique enquetant sur une affaire de contamination de l'eau ?", 'da': 'I hvilken film spiller Julia Roberts en advokatassistent, der undersøger en vandforureningssag?'},
 'opt1': {'en': 'Erin Brockovich', 'de': 'Erin Brockovich', 'es': 'Erin Brockovich', 'nl': 'Erin Brockovich', 'sv': 'Erin Brockovich', 'fr': 'Erin Brockovich', 'da': 'Erin Brockovich'},
 'opt2': {'en': 'A Civil Action', 'de': 'Civil Action', 'es': 'Acción civil', 'nl': 'A Civil Action', 'sv': 'A Civil Action', 'fr': 'Un recours en justice', 'da': 'A Civil Action'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Hampi?', 'de': 'In welchem Land liegt die antike Stadt Hampi?', 'es': '¿En que pais esta la antigua ciudad de Hampi?', 'nl': 'In welk land ligt de oude stad Hampi?', 'sv': 'I vilket land ligger den antika staden Hampi?', 'fr': "Dans quel pays se trouve l'ancienne cite de Hampi ?", 'da': 'Hvilket land huser den antikke by Hampi?'},
 'opt1': {'en': 'Sri Lanka', 'de': 'Sri Lanka', 'es': 'Sri Lanka', 'nl': 'Sri Lanka', 'sv': 'Sri Lanka', 'fr': 'le Sri Lanka', 'da': 'Sri Lanka'},
 'opt2': {'en': 'India', 'de': 'Indien', 'es': 'India', 'nl': 'India', 'sv': 'Indien', 'fr': "l'Inde", 'da': 'Indien'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which type of skirt, popular in the 1950s, is very full and worn with a petticoat?', 'de': 'Welche Art von Rock, populaer in den 1950er Jahren, ist sehr weit und wird mit einem Unterrock getragen?', 'es': '¿Que tipo de falda, popular en los años 50, es muy amplia y se lleva con enagua?', 'nl': 'Welk type rok, populair in de jaren 50, is erg wijd en wordt gedragen met een onderrok?', 'sv': 'Vilken typ av kjol, popular pa 1950-talet, ar mycket full och baers med underkjol?', 'fr': "Quel type de jupe, populaire dans les annees 1950, est tres ample et se porte avec un jupon ?", 'da': 'Hvilken type nederdel, populær i 1950’erne, er meget fuld og bæres med underskørt?'},
 'opt1': {'en': 'Pencil skirt', 'de': 'Bleistiftrock', 'es': 'Falda lapiz', 'nl': 'Potloodrok', 'sv': 'Pennkjol', 'fr': 'Jupe crayon', 'da': 'Blyantnederdel'},
 'opt2': {'en': 'Poodle skirt', 'de': 'Poodle Skirt', 'es': 'Falda poodle', 'nl': 'Poodle skirt', 'sv': 'Poodle skirt', 'fr': 'Jupe caniche', 'da': 'Poodle skirt'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which large flightless bird native to New Zealand is nocturnal and has no wings visible?', 'de': 'Welcher grosse flugunfaehige Vogel aus Neuseeland ist nachtaktiv und hat keine sichtbaren Fluegel?', 'es': '¿Que gran ave no voladora nativa de Nueva Zelanda es nocturna y no tiene alas visibles?', 'nl': 'Welke grote vliegloze vogel, inheems in Nieuw-Zeeland, is nachtactief en heeft geen zichtbare vleugels?', 'sv': 'Vilken stor flygofor fagel, inhemsk i Nya Zeeland, ar nattaktiv och har inga synliga vingar?', 'fr': "Quel grand oiseau incapable de voler, originaire de Nouvelle-Zelande, est nocturne et n'a pas d'ailes visibles ?", 'da': 'Hvilken stor flyveudygtig fugl, hjemmehørende i New Zealand, er natlig og har ingen synlige vinger?'},
 'opt1': {'en': 'Emu', 'de': 'Emu', 'es': 'Emu', 'nl': 'Emoe', 'sv': 'Emu', 'fr': 'Emeu', 'da': 'Emu'},
 'opt2': {'en': 'Kiwi', 'de': 'Kiwi', 'es': 'Kiwi', 'nl': 'Kiwi', 'sv': 'Kiwi', 'fr': 'Kiwi', 'da': 'Kiwi'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which First Lady was the first to hold a press conference exclusively for female reporters?', 'de': 'Welche First Lady hielt als erste eine Pressekonferenz ausschliesslich fuer weibliche Reporterinnen ab?', 'es': '¿Que primera dama fue la primera en celebrar una conferencia de prensa exclusiva para reporteras?', 'nl': 'Welke First Lady hield als eerste een persconferentie exclusief voor vrouwelijke journalisten?', 'sv': 'Vilken First Lady var forst med att halla en presskonferens uteslutande for kvinnliga reportrar?', 'fr': "Quelle Première dame a ete la premiere a tenir une conference de presse exclusivement pour les journalistes femmes ?", 'da': 'Hvilken førstedame var den første til at afholde en pressekonference udelukkende for kvindelige journalister?'},
 'opt1': {'en': 'Eleanor Roosevelt', 'de': 'Eleanor Roosevelt', 'es': 'Eleanor Roosevelt', 'nl': 'Eleanor Roosevelt', 'sv': 'Eleanor Roosevelt', 'fr': 'Eleanor Roosevelt', 'da': 'Eleanor Roosevelt'},
 'opt2': {'en': 'Lou Hoover', 'de': 'Lou Hoover', 'es': 'Lou Hoover', 'nl': 'Lou Hoover', 'sv': 'Lou Hoover', 'fr': 'Lou Hoover', 'da': 'Lou Hoover'}},

{'category': 'History', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which First Lady was a trained painter whose watercolors were exhibited publicly?', 'de': 'Welche First Lady war eine ausgebildete Malerin, deren Aquarelle oeffentlich ausgestellt wurden?', 'es': '¿Que primera dama fue una pintora entrenada cuyas acuarelas se exhibieron publicamente?', 'nl': 'Welke First Lady was een opgeleide schilder wier aquarellen publiekelijk werden tentoongesteld?', 'sv': 'Vilken First Lady var en utbildad malare vars akvareller stalldes ut offentligt?', 'fr': "Quelle Première dame etait une peintre formee dont les aquarelles ont ete exposees publiquement ?", 'da': 'Hvilken førstedame var en uddannet maler, hvis akvareller blev udstillet offentligt?'},
 'opt1': {'en': 'Eleanor Roosevelt', 'de': 'Eleanor Roosevelt', 'es': 'Eleanor Roosevelt', 'nl': 'Eleanor Roosevelt', 'sv': 'Eleanor Roosevelt', 'fr': 'Eleanor Roosevelt', 'da': 'Eleanor Roosevelt'},
 'opt2': {'en': 'Grace Coolidge', 'de': 'Grace Coolidge', 'es': 'Grace Coolidge', 'nl': 'Grace Coolidge', 'sv': 'Grace Coolidge', 'fr': 'Grace Coolidge', 'da': 'Grace Coolidge'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which Japanese author wrote 'Norwegian Wood'?", 'de': "Welcher japanische Autor schrieb 'Naokos Laecheln'?", 'es': "¿Que autor japones escribio 'Tokio Blues (Norwegian Wood)'?", 'nl': "Welke Japanse auteur schreef 'Norwegian Wood'?", 'sv': "Vilken japansk forfattare skrev 'Norwegian Wood'?", 'fr': "Quel auteur japonais a ecrit 'La Ballade de l'impossible (Norwegian Wood)' ?", 'da': "Hvilken japansk forfatter skrev 'Norwegian Wood'?"},
 'opt1': {'en': 'Yukio Mishima', 'de': 'Yukio Mishima', 'es': 'Yukio Mishima', 'nl': 'Yukio Mishima', 'sv': 'Yukio Mishima', 'fr': 'Yukio Mishima', 'da': 'Yukio Mishima'},
 'opt2': {'en': 'Haruki Murakami', 'de': 'Haruki Murakami', 'es': 'Haruki Murakami', 'nl': 'Haruki Murakami', 'sv': 'Haruki Murakami', 'fr': 'Haruki Murakami', 'da': 'Haruki Murakami'}},

{'category': 'Science', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which vitamin deficiency causes beriberi, affecting the nervous system?', 'de': 'Welcher Vitaminmangel verursacht Beriberi, das das Nervensystem betrifft?', 'es': '¿Que deficiencia vitaminica causa el beriberi, que afecta al sistema nervioso?', 'nl': 'Welk vitaminetekort veroorzaakt beriberi, dat het zenuwstelsel aantast?', 'sv': 'Vilken vitaminbrist orsakar beriberi, som paverkar nervsystemet?', 'fr': "Quelle carence en vitamine cause le beriberi, affectant le systeme nerveux ?", 'da': 'Hvilken vitaminmangel forårsager beriberi, som påvirker nervesystemet?'},
 'opt1': {'en': 'Vitamin B1 (Thiamine)', 'de': 'Vitamin B1 (Thiamin)', 'es': 'Vitamina B1 (Tiamina)', 'nl': 'Vitamine B1 (Thiamine)', 'sv': 'Vitamin B1 (Tiamin)', 'fr': 'Vitamine B1 (Thiamine)', 'da': 'Vitamin B1 (Thiamin)'},
 'opt2': {'en': 'Vitamin B12', 'de': 'Vitamin B12', 'es': 'Vitamina B12', 'nl': 'Vitamine B12', 'sv': 'Vitamin B12', 'fr': 'Vitamine B12', 'da': 'Vitamin B12'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which Mexican muralist painted large political works alongside his wife Frida Kahlo?', 'de': 'Welcher mexikanische Wandmaler malte grosse politische Werke neben seiner Frau Frida Kahlo?', 'es': '¿Que muralista mexicano pinto grandes obras politicas junto a su esposa Frida Kahlo?', 'nl': 'Welke Mexicaanse muurschilder maakte grote politieke werken naast zijn vrouw Frida Kahlo?', 'sv': 'Vilken mexikansk muralmalare malade stora politiska verk vid sidan av sin fru Frida Kahlo?', 'fr': "Quel muraliste mexicain a peint de grandes œuvres politiques aux cotes de sa femme Frida Kahlo ?", 'da': 'Hvilken mexicansk vægmaler malede store politiske værker ved siden af sin kone Frida Kahlo?'},
 'opt1': {'en': 'David Alfaro Siqueiros', 'de': 'David Alfaro Siqueiros', 'es': 'David Alfaro Siqueiros', 'nl': 'David Alfaro Siqueiros', 'sv': 'David Alfaro Siqueiros', 'fr': 'David Alfaro Siqueiros', 'da': 'David Alfaro Siqueiros'},
 'opt2': {'en': 'Diego Rivera', 'de': 'Diego Rivera', 'es': 'Diego Rivera', 'nl': 'Diego Rivera', 'sv': 'Diego Rivera', 'fr': 'Diego Rivera', 'da': 'Diego Rivera'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which African country is home to the ancient rock-hewn churches of Lalibela?', 'de': 'In welchem afrikanischen Land befinden sich die antiken, in Fels gehauenen Kirchen von Lalibela?', 'es': '¿Que pais africano alberga las antiguas iglesias talladas en roca de Lalibela?', 'nl': 'Welk Afrikaans land herbergt de oude uit rots gehouwen kerken van Lalibela?', 'sv': 'Vilket afrikanskt land huserar de antika klippkyrkorna i Lalibela?', 'fr': "Quel pays africain abrite les anciennes eglises rupestres de Lalibela ?", 'da': 'Hvilket afrikansk land huser de antikke klippehuggede kirker i Lalibela?'},
 'opt1': {'en': 'Ethiopia', 'de': 'Äthiopien', 'es': 'Etiopia', 'nl': 'Ethiopie', 'sv': 'Etiopien', 'fr': "l'Ethiopie", 'da': 'Etiopien'},
 'opt2': {'en': 'Eritrea', 'de': 'Eritrea', 'es': 'Eritrea', 'nl': 'Eritrea', 'sv': 'Eritrea', 'fr': "l'Erythree", 'da': 'Eritrea'}},

{'category': 'Music', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which French composer wrote 'Clair de Lune'?", 'de': "Welcher franzoesische Komponist schrieb 'Clair de Lune'?", 'es': "¿Que compositor frances escribio 'Claro de luna'?", 'nl': "Welke Franse componist schreef 'Clair de Lune'?", 'sv': "Vilken fransk kompositor skrev 'Clair de Lune'?", 'fr': "Quel compositeur français a ecrit 'Clair de Lune' ?", 'da': "Hvilken fransk komponist skrev 'Clair de Lune'?"},
 'opt1': {'en': 'Maurice Ravel', 'de': 'Maurice Ravel', 'es': 'Maurice Ravel', 'nl': 'Maurice Ravel', 'sv': 'Maurice Ravel', 'fr': 'Maurice Ravel', 'da': 'Maurice Ravel'},
 'opt2': {'en': 'Claude Debussy', 'de': 'Claude Debussy', 'es': 'Claude Debussy', 'nl': 'Claude Debussy', 'sv': 'Claude Debussy', 'fr': 'Claude Debussy', 'da': 'Claude Debussy'}},

{'category': 'Food', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Portuguese dish is salted, dried codfish prepared many different ways?', 'de': 'Welches portugiesische Gericht ist gesalzener, getrockneter Kabeljau, auf viele verschiedene Arten zubereitet?', 'es': '¿Que plato portugues es bacalao salado y seco preparado de muchas formas diferentes?', 'nl': 'Welk Portugees gerecht is gezouten, gedroogde kabeljauw op vele verschillende manieren bereid?', 'sv': 'Vilken portugisisk ratt ar saltad, torkad torsk tillagad pa manga olika satt?', 'fr': "Quel plat portugais est de la morue salee et sechee preparee de nombreuses façons differentes ?", 'da': 'Hvilken portugisisk ret er saltet, tørret torsk tilberedt på mange forskellige måder?'},
 'opt1': {'en': 'Bacalhau', 'de': 'Bacalhau', 'es': 'Bacalao', 'nl': 'Bacalhau', 'sv': 'Bacalhau', 'fr': 'Bacalhau', 'da': 'Bacalhau'},
 'opt2': {'en': 'Caldeirada', 'de': 'Caldeirada', 'es': 'Caldeirada', 'nl': 'Caldeirada', 'sv': 'Caldeirada', 'fr': 'Caldeirada', 'da': 'Caldeirada'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'In Hindu mythology, which god is known as the destroyer, part of the trinity with Brahma and Vishnu?', 'de': 'In der hinduistischen Mythologie, welcher Gott ist als Zerstoerer bekannt, Teil der Trinitat mit Brahma und Vishnu?', 'es': '¿En la mitologia hindu, que dios es conocido como el destructor, parte de la trinidad con Brahma y Vishnu?', 'nl': 'In de hindoemythologie, welke god staat bekend als de vernietiger, onderdeel van de drie-eenheid met Brahma en Vishnu?', 'sv': 'I hinduisk mytologi, vilken gud ar kand som forstoraren, del av treenigheten med Brahma och Vishnu?', 'fr': "Dans la mythologie hindoue, quel dieu est connu comme le destructeur, faisant partie de la trinite avec Brahma et Vishnu ?", 'da': 'I hinduistisk mytologi, hvilken gud er kendt som ødelæggeren, del af treenigheden med Brahma og Vishnu?'},
 'opt1': {'en': 'Indra', 'de': 'Indra', 'es': 'Indra', 'nl': 'Indra', 'sv': 'Indra', 'fr': 'Indra', 'da': 'Indra'},
 'opt2': {'en': 'Shiva', 'de': 'Shiva', 'es': 'Shiva', 'nl': 'Shiva', 'sv': 'Shiva', 'fr': 'Shiva', 'da': 'Shiva'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which actress won an Oscar for playing a maid in 'The Help'?", 'de': "Welche Schauspielerin gewann einen Oscar fuer ihre Rolle als Dienstmaedchen in 'The Help'?", 'es': "¿Que actriz gano un Oscar por interpretar a una criada en 'Criadas y señoras'?", 'nl': "Welke actrice won een Oscar voor haar rol als dienstmeisje in 'The Help'?", 'sv': "Vilken skadespelerska vann en Oscar for att ha spelat en hushallerska i 'The Help'?", 'fr': "Quelle actrice a remporte un Oscar pour son role de domestique dans 'La Couleur des sentiments' ?", 'da': "Hvilken skuespiller vandt en Oscar for at spille en tjenestepige i 'The Help'?"},
 'opt1': {'en': 'Octavia Spencer', 'de': 'Octavia Spencer', 'es': 'Octavia Spencer', 'nl': 'Octavia Spencer', 'sv': 'Octavia Spencer', 'fr': 'Octavia Spencer', 'da': 'Octavia Spencer'},
 'opt2': {'en': 'Viola Davis', 'de': 'Viola Davis', 'es': 'Viola Davis', 'nl': 'Viola Davis', 'sv': 'Viola Davis', 'fr': 'Viola Davis', 'da': 'Viola Davis'}},

{'category': 'Space', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which spacecraft was the first to land a rover on Mars in 1997?', 'de': 'Welches Raumfahrzeug landete 1997 als erstes einen Rover auf dem Mars?', 'es': '¿Que nave espacial fue la primera en aterrizar un rover en Marte en 1997?', 'nl': 'Welk ruimtevaartuig landde in 1997 als eerste een rover op Mars?', 'sv': 'Vilket rymdfarkost var forst med att landa en rover pa Mars 1997?', 'fr': "Quel vaisseau spatial a ete le premier a poser un rover sur Mars en 1997 ?", 'da': 'Hvilket rumfartøj var det første til at lande en rover på Mars i 1997?'},
 'opt1': {'en': 'Viking 1', 'de': 'Viking 1', 'es': 'Viking 1', 'nl': 'Viking 1', 'sv': 'Viking 1', 'fr': 'Viking 1', 'da': 'Viking 1'},
 'opt2': {'en': 'Mars Pathfinder', 'de': 'Mars Pathfinder', 'es': 'Mars Pathfinder', 'nl': 'Mars Pathfinder', 'sv': 'Mars Pathfinder', 'fr': 'Mars Pathfinder', 'da': 'Mars Pathfinder'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which South American country's capital, Quito, is one of the highest capital cities in the world?", 'de': 'Welches suedamerikanische Land hat mit Quito eine der hoechstgelegenen Hauptstaedte der Welt?', 'es': '¿Que pais sudamericano tiene como capital a Quito, una de las capitales mas altas del mundo?', 'nl': 'Welk Zuid-Amerikaans land heeft met Quito een van de hoogst gelegen hoofdsteden ter wereld?', 'sv': 'Vilket sydamerikanskt land har Quito som huvudstad, en av varldens hogst belagna huvudstader?', 'fr': "Quel pays sud-americain a pour capitale Quito, l'une des capitales les plus hautes du monde ?", 'da': 'Hvilket sydamerikansk land har Quito som hovedstad, en af verdens højest beliggende hovedstæder?'},
 'opt1': {'en': 'Bolivia', 'de': 'Bolivien', 'es': 'Bolivia', 'nl': 'Bolivia', 'sv': 'Bolivia', 'fr': 'la Bolivie', 'da': 'Bolivia'},
 'opt2': {'en': 'Ecuador', 'de': 'Ecuador', 'es': 'Ecuador', 'nl': 'Ecuador', 'sv': 'Ecuador', 'fr': "l'Equateur", 'da': 'Ecuador'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Italian designer founded a house known for its bamboo-handled handbags?', 'de': 'Welcher italienische Designer gruendete ein Haus, das fuer Handtaschen mit Bambusgriff bekannt ist?', 'es': '¿Que diseñador italiano fundo una casa conocida por sus bolsos con asa de bambu?', 'nl': 'Welke Italiaanse ontwerper richtte een huis op dat bekend staat om zijn handtassen met bamboehengsel?', 'sv': 'Vilken italiensk designer grundade ett hus kant for sina handvaskor med bambuhandtag?', 'fr': "Quel createur italien a fonde une maison connue pour ses sacs a main a poignee de bambou ?", 'da': 'Hvilken italiensk designer grundlagde et hus kendt for sine håndtasker med bambushank?'},
 'opt1': {'en': 'Gucci', 'de': 'Gucci', 'es': 'Gucci', 'nl': 'Gucci', 'sv': 'Gucci', 'fr': 'Gucci', 'da': 'Gucci'},
 'opt2': {'en': 'Prada', 'de': 'Prada', 'es': 'Prada', 'nl': 'Prada', 'sv': 'Prada', 'fr': 'Prada', 'da': 'Prada'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which ancient Chinese philosopher is credited with founding Taoism?', 'de': 'Welcher antike chinesische Philosoph gilt als Begruender des Taoismus?', 'es': '¿Que antiguo filosofo chino es reconocido por fundar el taoismo?', 'nl': 'Welke oude Chinese filosoof wordt gecrediteerd met het stichten van het taoisme?', 'sv': 'Vilken antik kinesisk filosof tillskrivs att ha grundat taoismen?', 'fr': "Quel philosophe chinois antique est credite d'avoir fonde le taoisme ?", 'da': 'Hvilken oldgammel kinesisk filosof krediteres for at have grundlagt taoismen?'},
 'opt1': {'en': 'Laozi', 'de': 'Laozi', 'es': 'Laozi', 'nl': 'Laozi', 'sv': 'Laozi', 'fr': 'Laozi', 'da': 'Laozi'},
 'opt2': {'en': 'Confucius', 'de': 'Konfuzius', 'es': 'Confucio', 'nl': 'Confucius', 'sv': 'Konfucius', 'fr': 'Confucius', 'da': 'Konfucius'}},

{'category': 'History', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which quilting pattern features a central medallion surrounded by concentric borders?', 'de': 'Welches Quiltmuster zeigt ein zentrales Medaillon, umgeben von konzentrischen Rahmen?', 'es': '¿Que patron de acolchado presenta un medallon central rodeado de bordes concentricos?', 'nl': 'Welk quiltpatroon toont een centraal medaillon omgeven door concentrische randen?', 'sv': 'Vilket lapptacksmonster har en central medaljong omgiven av koncentriska kanter?', 'fr': "Quel motif de patchwork presente un medaillon central entoure de bordures concentriques ?", 'da': 'Hvilket quiltmønster har en central medaljon omgivet af koncentriske kanter?'},
 'opt1': {'en': 'Sampler quilt', 'de': 'Sampler-Quilt', 'es': 'Edredon muestrario', 'nl': 'Samplerquilt', 'sv': 'Samplerquilt', 'fr': 'Patchwork echantillon', 'da': 'Samplerquilt'},
 'opt2': {'en': 'Medallion quilt', 'de': 'Medaillon-Quilt', 'es': 'Edredon medallon', 'nl': 'Medaillonquilt', 'sv': 'Medaljongquilt', 'fr': 'Patchwork medaillon', 'da': 'Medaljonquilt'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which flowering tree is famous for blooming pink in Washington D.C. each spring, a gift from Japan?', 'de': 'Welcher bluehende Baum ist beruehmt fuer seine rosa Blueten in Washington D.C. jeden Fruehling, ein Geschenk aus Japan?', 'es': '¿Que arbol en flor es famoso por florecer en rosa en Washington D.C. cada primavera, un regalo de Japon?', 'nl': 'Welke bloeiende boom staat bekend om zijn roze bloesems in Washington D.C. elk voorjaar, een geschenk uit Japan?', 'sv': 'Vilket blommande trad ar kant for att blomma rosa i Washington D.C. varje var, en gava fran Japan?', 'fr': "Quel arbre a fleurs est celebre pour ses fleurs roses a Washington D.C. chaque printemps, un cadeau du Japon ?", 'da': 'Hvilket blomstrende træ er berømt for at blomstre pink i Washington D.C. hvert forår, en gave fra Japan?'},
 'opt1': {'en': 'Cherry blossom', 'de': 'Kirschbluete', 'es': 'Flor de cerezo', 'nl': 'Kersenbloesem', 'sv': 'Korsbarsblom', 'fr': 'Fleur de cerisier', 'da': 'Kirsebærblomst'},
 'opt2': {'en': 'Dogwood', 'de': 'Hartriegel', 'es': 'Cornejo', 'nl': 'Kornoelje', 'sv': 'Kornell', 'fr': 'Cornouiller', 'da': 'Kornel'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which architect designed the Guggenheim Museum Bilbao with its titanium curves?', 'de': 'Welcher Architekt entwarf das Guggenheim-Museum Bilbao mit seinen Titankurven?', 'es': '¿Que arquitecto diseño el Museo Guggenheim de Bilbao con sus curvas de titanio?', 'nl': 'Welke architect ontwierp het Guggenheim Museum Bilbao met zijn titanium curven?', 'sv': 'Vilken arkitekt designade Guggenheimmuseet i Bilbao med sina titankurvor?', 'fr': "Quel architecte a conçu le musee Guggenheim de Bilbao avec ses courbes en titane ?", 'da': 'Hvilken arkitekt designede Guggenheim-museet i Bilbao med dets titaniumkurver?'},
 'opt1': {'en': 'Frank Gehry', 'de': 'Frank Gehry', 'es': 'Frank Gehry', 'nl': 'Frank Gehry', 'sv': 'Frank Gehry', 'fr': 'Frank Gehry', 'da': 'Frank Gehry'},
 'opt2': {'en': 'Zaha Hadid', 'de': 'Zaha Hadid', 'es': 'Zaha Hadid', 'nl': 'Zaha Hadid', 'sv': 'Zaha Hadid', 'fr': 'Zaha Hadid', 'da': 'Zaha Hadid'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'In Hindu mythology, which goddess is associated with knowledge and the arts?', 'de': 'In der hinduistischen Mythologie, welche Goettin wird mit Wissen und Kunst assoziiert?', 'es': '¿En la mitologia hindu, que diosa se asocia con el conocimiento y las artes?', 'nl': 'In de hindoemythologie, welke godin wordt geassocieerd met kennis en de kunsten?', 'sv': 'I hinduisk mytologi, vilken gudinna forknippas med kunskap och konsterna?', 'fr': "Dans la mythologie hindoue, quelle deesse est associee a la connaissance et aux arts ?", 'da': 'I hinduistisk mytologi, hvilken gudinde forbindes med viden og kunsten?'},
 'opt1': {'en': 'Parvati', 'de': 'Parvati', 'es': 'Parvati', 'nl': 'Parvati', 'sv': 'Parvati', 'fr': 'Parvati', 'da': 'Parvati'},
 'opt2': {'en': 'Saraswati', 'de': 'Saraswati', 'es': 'Saraswati', 'nl': 'Saraswati', 'sv': 'Saraswati', 'fr': 'Saraswati', 'da': 'Saraswati'}},
]
