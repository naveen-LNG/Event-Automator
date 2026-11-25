# -*- coding: utf-8 -*-
# Trivia Challenge batch #11. See trivia_questions_batch_01.py for schema notes.

QUESTIONS = [
{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which white grain is ground into flour to make bread?', 'de': 'Welches weisse Getreide wird zu Mehl gemahlen, um Brot zu backen?', 'es': '¿Que grano blanco se muele en harina para hacer pan?', 'nl': 'Welke witte graansoort wordt gemalen tot meel om brood te maken?', 'sv': 'Vilket vitt sadesslag mals till mjol for att gora brod?', 'fr': "Quelle cereale blanche est moulue en farine pour faire du pain ?", 'da': 'Hvilken hvid kerne males til mel for at lave brød?'},
 'opt1': {'en': 'Wheat', 'de': 'Weizen', 'es': 'Trigo', 'nl': 'Tarwe', 'sv': 'Vete', 'fr': 'Le ble', 'da': 'Hvede'},
 'opt2': {'en': 'Rice', 'de': 'Reis', 'es': 'Arroz', 'nl': 'Rijst', 'sv': 'Ris', 'fr': 'Le riz', 'da': 'Ris'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Delhi?', 'de': 'In welchem Land liegt die antike Stadt Delhi?', 'es': '¿En que pais esta la antigua ciudad de Delhi?', 'nl': 'In welk land ligt de oude stad Delhi?', 'sv': 'I vilket land ligger den antika staden Delhi?', 'fr': "Dans quel pays se trouve l'ancienne cite de Delhi ?", 'da': 'Hvilket land huser den antikke by Delhi?'},
 'opt1': {'en': 'India', 'de': 'Indien', 'es': 'India', 'nl': 'India', 'sv': 'Indien', 'fr': "l'Inde", 'da': 'Indien'},
 'opt2': {'en': 'Nepal', 'de': 'Nepal', 'es': 'Nepal', 'nl': 'Nepal', 'sv': 'Nepal', 'fr': 'le Nepal', 'da': 'Nepal'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a boy who travels back in time in a DeLorean car?', 'de': 'In welchem Filmklassiker reist ein Junge in einem DeLorean-Auto in die Vergangenheit?', 'es': '¿Que pelicula clasica presenta a un joven que viaja al pasado en un auto DeLorean?', 'nl': 'In welke klassieke film reist een jongen terug in de tijd in een DeLorean-auto?', 'sv': 'I vilken klassisk film reser en pojke tillbaka i tiden i en DeLorean-bil?', 'fr': "Quel film classique met en scene un garçon qui voyage dans le temps dans une voiture DeLorean ?", 'da': 'Hvilken klassisk film har en dreng, der rejser tilbage i tiden i en DeLorean-bil?'},
 'opt1': {'en': 'Back to the Future', 'de': 'Zurueck in die Zukunft', 'es': 'Volver al futuro', 'nl': 'Back to the Future', 'sv': 'Tillbaka till framtiden', 'fr': 'Retour vers le futur', 'da': 'Tilbage til fremtiden'},
 'opt2': {'en': 'The Time Machine', 'de': 'Die Zeitmaschine', 'es': 'La maquina del tiempo', 'nl': 'The Time Machine', 'sv': 'Tidsmaskinen', 'fr': 'La Machine a explorer le temps', 'da': 'Tidsmaskinen'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which pet is known for chirping and is often kept in a cage?', 'de': 'Welches Haustier zwitschert und wird oft in einem Kaefig gehalten?', 'es': '¿Que mascota es conocida por trinar y se mantiene a menudo en una jaula?', 'nl': 'Welk huisdier staat bekend om tjilpen en wordt vaak in een kooi gehouden?', 'sv': 'Vilket sallskapsdjur kvittrar och halls ofta i en bur?', 'fr': "Quel animal de compagnie gazouille et est souvent garde en cage ?", 'da': 'Hvilket kæledyr kvidrer og holdes ofte i et bur?'},
 'opt1': {'en': 'Hamster', 'de': 'Hamster', 'es': 'Hamster', 'nl': 'Hamster', 'sv': 'Hamster', 'fr': 'Hamster', 'da': 'Hamster'},
 'opt2': {'en': 'Bird (canary)', 'de': 'Vogel (Kanarienvogel)', 'es': 'Ave (canario)', 'nl': 'Vogel (kanarie)', 'sv': 'Fagel (kanariefagel)', 'fr': 'Oiseau (canari)', 'da': 'Fugl (kanariefugl)'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which yellow spread is made from churned cream and used on toast?', 'de': 'Welcher gelbe Aufstrich wird aus gebutterter Sahne hergestellt und auf Toast verwendet?', 'es': '¿Que untable amarillo se hace de crema batida y se usa en las tostadas?', 'nl': 'Welke gele smeersel wordt gemaakt van gekarnde room en gebruikt op toast?', 'sv': 'Vilket gult palagg gors av kärnad gradde och anvands pa rostat brod?', 'fr': "Quelle tartinade jaune est faite de creme baratee et utilisee sur les toasts ?", 'da': 'Hvilket gult pålæg er lavet af kærnet fløde og bruges på ristet brød?'},
 'opt1': {'en': 'Margarine', 'de': 'Margarine', 'es': 'Margarina', 'nl': 'Margarine', 'sv': 'Margarin', 'fr': 'La margarine', 'da': 'Margarine'},
 'opt2': {'en': 'Butter', 'de': 'Butter', 'es': 'Mantequilla', 'nl': 'Boter', 'sv': 'Smor', 'fr': 'Le beurre', 'da': 'Smør'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Bangkok?', 'de': 'In welchem Land liegt die antike Stadt Bangkok?', 'es': '¿En que pais esta la antigua ciudad de Bangkok?', 'nl': 'In welk land ligt de oude stad Bangkok?', 'sv': 'I vilket land ligger den antika staden Bangkok?', 'fr': "Dans quel pays se trouve l'ancienne cite de Bangkok ?", 'da': 'Hvilket land huser den antikke by Bangkok?'},
 'opt1': {'en': 'Vietnam', 'de': 'Vietnam', 'es': 'Vietnam', 'nl': 'Vietnam', 'sv': 'Vietnam', 'fr': 'le Vietnam', 'da': 'Vietnam'},
 'opt2': {'en': 'Thailand', 'de': 'Thailand', 'es': 'Tailandia', 'nl': 'Thailand', 'sv': 'Thailand', 'fr': 'la Thailande', 'da': 'Thailand'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. president is known for the Louisiana Purchase?', 'de': 'Welcher US-Praesident ist fuer den Louisiana-Kauf bekannt?', 'es': '¿Que presidente de EE.UU. es conocido por la Compra de Luisiana?', 'nl': 'Welke Amerikaanse president staat bekend om de Louisiana Purchase?', 'sv': 'Vilken amerikansk president ar kand for Louisianakopet?', 'fr': "Quel president americain est connu pour l'achat de la Louisiane ?", 'da': 'Hvilken amerikansk præsident er kendt for Louisiana-købet?'},
 'opt1': {'en': 'James Madison', 'de': 'James Madison', 'es': 'James Madison', 'nl': 'James Madison', 'sv': 'James Madison', 'fr': 'James Madison', 'da': 'James Madison'},
 'opt2': {'en': 'Thomas Jefferson', 'de': 'Thomas Jefferson', 'es': 'Thomas Jefferson', 'nl': 'Thomas Jefferson', 'sv': 'Thomas Jefferson', 'fr': 'Thomas Jefferson', 'da': 'Thomas Jefferson'}},

{'category': 'Arts', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which craft involves shaping metal wire into jewelry or sculptures?', 'de': 'Bei welchem Handwerk wird Metalldraht zu Schmuck oder Skulpturen geformt?', 'es': '¿Que manualidad implica dar forma a alambre metalico en joyeria o esculturas?', 'nl': 'Bij welk handwerk vorm je metaaldraad tot sieraden of sculpturen?', 'sv': 'Vilket hantverk innebar att forma metalltrad till smycken eller skulpturer?', 'fr': "Quel art manuel consiste a façonner du fil metallique en bijoux ou sculptures ?", 'da': 'Hvilket håndværk involverer at forme metaltråd til smykker eller skulpturer?'},
 'opt1': {'en': 'Wire wrapping', 'de': 'Drahtwickeltechnik', 'es': 'Enrollado de alambre', 'nl': 'Draadwikkelen', 'sv': 'Tradlindning', 'fr': 'Le wire wrapping', 'da': 'Trådvikling'},
 'opt2': {'en': 'Metalsmithing', 'de': 'Metallschmieden', 'es': 'Metalisteria', 'nl': 'Metaalbewerking', 'sv': 'Metallsmide', 'fr': 'La ferronnerie', 'da': 'Metalsmedning'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which woodwind instrument has a double reed and a deep, rich tone?', 'de': 'Welches Holzblasinstrument hat ein Doppelrohrblatt und einen tiefen, satten Klang?', 'es': '¿Que instrumento de viento madera tiene doble lengueta y un tono profundo y rico?', 'nl': 'Welk houtblazersinstrument heeft een dubbel riet en een diepe, rijke toon?', 'sv': 'Vilket trablasinstrument har dubbelt rorblad och en djup, rik ton?', 'fr': "Quel instrument a vent en bois a une anche double et une tonalite profonde et riche ?", 'da': 'Hvilket træblæserinstrument har et dobbelt rørblad og en dyb, rig tone?'},
 'opt1': {'en': 'Bassoon', 'de': 'Fagott', 'es': 'Fagot', 'nl': 'Fagot', 'sv': 'Fagott', 'fr': 'Basson', 'da': 'Fagot'},
 'opt2': {'en': 'Oboe', 'de': 'Oboe', 'es': 'Oboe', 'nl': 'Hobo', 'sv': 'Oboe', 'fr': 'Hautbois', 'da': 'Obo'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which natural event involves a volcano releasing lava and ash?', 'de': 'Welches Naturereignis beinhaltet, dass ein Vulkan Lava und Asche freisetzt?', 'es': '¿Que evento natural implica que un volcan libere lava y ceniza?', 'nl': 'Welke natuurlijke gebeurtenis omvat een vulkaan die lava en as uitstoot?', 'sv': 'Vilken naturhandelse innebar att en vulkan slapper ut lava och aska?', 'fr': "Quel evenement naturel implique un volcan liberant lave et cendres ?", 'da': 'Hvilken naturbegivenhed involverer en vulkan, der udsender lava og aske?'},
 'opt1': {'en': 'Earthquake', 'de': 'Erdbeben', 'es': 'Terremoto', 'nl': 'Aardbeving', 'sv': 'Jordbavning', 'fr': 'Tremblement de terre', 'da': 'Jordskælv'},
 'opt2': {'en': 'Volcanic eruption', 'de': 'Vulkanausbruch', 'es': 'Erupcion volcanica', 'nl': 'Vulkaanuitbarsting', 'sv': 'Vulkanutbrott', 'fr': 'Eruption volcanique', 'da': 'Vulkanudbrud'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which classic film features a young girl who falls down a rabbit hole into a fantastical world?', 'de': 'In welchem Filmklassiker faellt ein junges Maedchen durch ein Kaninchenloch in eine fantastische Welt?', 'es': '¿Que pelicula clasica presenta a una niña que cae por una madriguera a un mundo fantastico?', 'nl': 'In welke klassieke film valt een klein meisje door een konijnenhol naar een fantastische wereld?', 'sv': 'I vilken klassisk film ramlar en liten flicka ner i ett kaninhal till en fantastisk varld?', 'fr': "Quel film classique met en scene une petite fille qui tombe dans un terrier vers un monde fantastique ?", 'da': 'Hvilken klassisk film har en lille pige, der falder ned i et kaninhul til en fantastisk verden?'},
 'opt1': {'en': 'Alice in Wonderland', 'de': 'Alice im Wunderland', 'es': 'Alicia en el pais de las maravillas', 'nl': 'Alice in Wonderland', 'sv': 'Alice i Underlandet', 'fr': "Alice au pays des merveilles", 'da': 'Alice i Eventyrland'},
 'opt2': {'en': 'Coraline', 'de': 'Coraline', 'es': 'Coraline', 'nl': 'Coraline', 'sv': 'Coraline', 'fr': 'Coraline', 'da': 'Coraline'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which U.S. state is known as the 'Ocean State'?", 'de': "Welcher US-Bundesstaat ist als 'Ocean State' bekannt?", 'es': "¿Que estado de EE.UU. es conocido como el 'Estado Oceanico'?", 'nl': "Welke Amerikaanse staat staat bekend als de 'Ocean State'?", 'sv': "Vilken amerikansk delstat ar kand som 'Ocean State'?", 'fr': "Quel État americain est connu comme le 'Ocean State' (État de l'ocean) ?", 'da': "Hvilken amerikansk delstat er kendt som 'Ocean State'?"},
 'opt1': {'en': 'Rhode Island', 'de': 'Rhode Island', 'es': 'Rhode Island', 'nl': 'Rhode Island', 'sv': 'Rhode Island', 'fr': 'le Rhode Island', 'da': 'Rhode Island'},
 'opt2': {'en': 'Delaware', 'de': 'Delaware', 'es': 'Delaware', 'nl': 'Delaware', 'sv': 'Delaware', 'fr': 'le Delaware', 'da': 'Delaware'}},

{'category': 'Food', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which brown legume is used to make traditional hummus?', 'de': 'Welche braune Huelsenfrucht wird fuer traditionellen Hummus verwendet?', 'es': '¿Que legumbre marron se usa para hacer el hummus tradicional?', 'nl': 'Welke bruine peulvrucht wordt gebruikt om traditionele hummus te maken?', 'sv': 'Vilken brun balja anvands for att gora traditionell hummus?', 'fr': "Quel legume brun est utilise pour faire le houmous traditionnel ?", 'da': 'Hvilken brun bælgfrugt bruges til at lave traditionel houmous?'},
 'opt1': {'en': 'Chickpea', 'de': 'Kichererbse', 'es': 'Garbanzo', 'nl': 'Kikkererwt', 'sv': 'Kikart', 'fr': 'Le pois chiche', 'da': 'Kikært'},
 'opt2': {'en': 'Lentil', 'de': 'Linse', 'es': 'Lenteja', 'nl': 'Linze', 'sv': 'Linsboenor', 'fr': 'La lentille', 'da': 'Linse'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which sport involves skiing down a mountain while shooting at targets?', 'de': 'Bei welcher Sportart faehrt man Ski einen Berg hinunter und schiesst dabei auf Ziele?', 'es': '¿Que deporte consiste en esquiar cuesta abajo mientras se dispara a blancos?', 'nl': 'Bij welke sport ski je een berg af terwijl je op doelen schiet?', 'sv': 'Vilken sport innebar att aka skidor nedfor ett berg medan man skjuter mot mal?', 'fr': "Quel sport consiste a skier en descente tout en tirant sur des cibles ?", 'da': 'Hvilken sport involverer at stå på ski ned ad et bjerg, mens man skyder mod mål?'},
 'opt1': {'en': 'Biathlon', 'de': 'Biathlon', 'es': 'Biatlon', 'nl': 'Biatlon', 'sv': 'Skidskytte', 'fr': 'Le biathlon', 'da': 'Skiskydning'},
 'opt2': {'en': 'Nordic combined', 'de': 'Nordische Kombination', 'es': 'Combinada nordica', 'nl': 'Noordse combinatie', 'sv': 'Nordisk kombination', 'fr': 'Le combine nordique', 'da': 'Nordisk kombination'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which large sea creature is known for its blowhole and is a mammal, not a fish?', 'de': 'Welches grosse Meereswesen ist fuer sein Blasloch bekannt und ist ein Saeugetier, kein Fisch?', 'es': '¿Que gran criatura marina es conocida por su espiraculo y es un mamifero, no un pez?', 'nl': 'Welk groot zeewezen staat bekend om zijn spuitgat en is een zoogdier, geen vis?', 'sv': 'Vilken stor havsvarelse ar kand for sitt blashal och ar ett daggdjur, inte en fisk?', 'fr': "Quelle grande creature marine est connue pour son event et est un mammifere, pas un poisson ?", 'da': 'Hvilket stort havvæsen er kendt for sit blæsehul og er et pattedyr, ikke en fisk?'},
 'opt1': {'en': 'Shark', 'de': 'Hai', 'es': 'Tiburon', 'nl': 'Haai', 'sv': 'Haj', 'fr': 'Requin', 'da': 'Haj'},
 'opt2': {'en': 'Whale', 'de': 'Wal', 'es': 'Ballena', 'nl': 'Walvis', 'sv': 'Val', 'fr': 'Baleine', 'da': 'Hval'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which holiday involves fasting from dawn to sunset for a month, observed by Muslims?', 'de': 'Bei welchem Feiertag fastet man einen Monat lang von der Morgendaemmerung bis zum Sonnenuntergang, beobachtet von Muslimen?', 'es': '¿Que festividad implica ayunar desde el amanecer hasta el atardecer durante un mes, observada por musulmanes?', 'nl': 'Bij welke feestdag vast je een maand lang van zonsopgang tot zonsondergang, gehouden door moslims?', 'sv': 'Vilken hogtid innebar att fasta fran gryning till solnedgang under en manad, firad av muslimer?', 'fr': "Quelle fete implique de jeuner de l'aube au coucher du soleil pendant un mois, observee par les musulmans ?", 'da': 'Hvilken helligdag involverer faste fra daggry til solnedgang i en måned, holdt af muslimer?'},
 'opt1': {'en': 'Yom Kippur', 'de': 'Jom Kippur', 'es': 'Yom Kipur', 'nl': 'Jom Kipoer', 'sv': 'Yom Kippur', 'fr': 'Yom Kippour', 'da': 'Yom Kippur'},
 'opt2': {'en': 'Ramadan', 'de': 'Ramadan', 'es': 'Ramadan', 'nl': 'Ramadan', 'sv': 'Ramadan', 'fr': 'Le ramadan', 'da': 'Ramadan'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which continent is home to Mount Everest?', 'de': 'Auf welchem Kontinent liegt der Mount Everest?', 'es': '¿En que continente esta el Monte Everest?', 'nl': 'Op welk continent ligt de Mount Everest?', 'sv': 'Pa vilken kontinent ligger Mount Everest?', 'fr': "Sur quel continent se trouve le mont Everest ?", 'da': 'Hvilket kontinent huser Mount Everest?'},
 'opt1': {'en': 'Europe', 'de': 'Europa', 'es': 'Europa', 'nl': 'Europa', 'sv': 'Europa', 'fr': "l'Europe", 'da': 'Europa'},
 'opt2': {'en': 'Asia', 'de': 'Asien', 'es': 'Asia', 'nl': 'Azie', 'sv': 'Asien', 'fr': "l'Asie", 'da': 'Asien'}},

{'category': 'Literature', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which author wrote 'Winnie-the-Pooh'?", 'de': "Welcher Autor schrieb 'Pu der Baer'?", 'es': "¿Que autor escribio 'Winnie the Pooh'?", 'nl': "Welke auteur schreef 'Winnie de Poeh'?", 'sv': "Vilken forfattare skrev 'Nalle Puh'?", 'fr': "Quel auteur a ecrit 'Winnie l'ourson' ?", 'da': "Hvilken forfatter skrev 'Peter Plys'?"},
 'opt1': {'en': 'A.A. Milne', 'de': 'A.A. Milne', 'es': 'A.A. Milne', 'nl': 'A.A. Milne', 'sv': 'A.A. Milne', 'fr': 'A.A. Milne', 'da': 'A.A. Milne'},
 'opt2': {'en': 'Kenneth Grahame', 'de': 'Kenneth Grahame', 'es': 'Kenneth Grahame', 'nl': 'Kenneth Grahame', 'sv': 'Kenneth Grahame', 'fr': 'Kenneth Grahame', 'da': 'Kenneth Grahame'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic toy is a colorful cube puzzle with rotating sides?', 'de': 'Welches klassische Spielzeug ist ein buntes Wuerfelpuzzle mit drehbaren Seiten?', 'es': '¿Que juguete clasico es un cubo rompecabezas colorido con lados giratorios?', 'nl': 'Welk klassiek speelgoed is een kleurrijke kubuspuzzel met draaiende zijden?', 'sv': 'Vilken klassisk leksak ar ett fargglatt kubpussel med roterande sidor?', 'fr': "Quel jouet classique est un casse-tete cubique colore avec des faces rotatives ?", 'da': 'Hvilket klassisk legetøj er en farverig kubepuslespil med roterende sider?'},
 'opt1': {'en': 'Puzzle box', 'de': 'Puzzlebox', 'es': 'Caja rompecabezas', 'nl': 'Puzzeldoos', 'sv': 'Pusseldosa', 'fr': 'La boite casse-tete', 'da': 'Puslekasse'},
 'opt2': {'en': "Rubik's Cube", 'de': 'Zauberwuerfel', 'es': 'Cubo de Rubik', 'nl': 'Rubiks kubus', 'sv': 'Rubiks kub', 'fr': 'Le Rubik\'s Cube', 'da': 'Rubiks terning'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which red spice comes from dried chili peppers and adds heat to dishes?', 'de': 'Welches rote Gewuerz stammt von getrockneten Chilischoten und verleiht Gerichten Schaerfe?', 'es': '¿Que especia roja proviene de chiles secos y da picante a los platos?', 'nl': 'Welke rode specerij komt van gedroogde chilipepers en geeft pittigheid aan gerechten?', 'sv': 'Vilken rod krydda kommer fran torkade chilipeppar och ger hetta at ratter?', 'fr': "Quelle epice rouge provient de piments seches et pimente les plats ?", 'da': 'Hvilket rødt krydderi kommer fra tørrede chilipeber og giver hede til retter?'},
 'opt1': {'en': 'Paprika', 'de': 'Paprika', 'es': 'Pimenton', 'nl': 'Paprikapoeder', 'sv': 'Paprikapulver', 'fr': 'Le paprika', 'da': 'Paprika'},
 'opt2': {'en': 'Chili powder', 'de': 'Chilipulver', 'es': 'Chile en polvo', 'nl': 'Chilipoeder', 'sv': 'Chilipulver', 'fr': 'La poudre de chili', 'da': 'Chilipulver'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. holiday celebrates mothers in May?', 'de': 'Welcher US-Feiertag im Mai feiert Muetter?', 'es': '¿Que festividad de EE.UU. en mayo celebra a las madres?', 'nl': 'Welke Amerikaanse feestdag in mei viert moeders?', 'sv': 'Vilken amerikansk hogtid i maj firar mammor?', 'fr': "Quelle fete americaine de mai celebre les meres ?", 'da': 'Hvilken amerikansk helligdag i maj fejrer mødre?'},
 'opt1': {'en': "Father's Day", 'de': 'Vatertag', 'es': 'Dia del Padre', 'nl': 'Vaderdag', 'sv': 'Fars dag', 'fr': 'La fete des Peres', 'da': 'Fars dag'},
 'opt2': {'en': "Mother's Day", 'de': 'Muttertag', 'es': 'Dia de la Madre', 'nl': 'Moederdag', 'sv': 'Mors dag', 'fr': "La fete des Meres", 'da': 'Mors dag'}},

{'category': 'Music', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which decade is associated with the rise of Motown music?', 'de': 'Mit welchem Jahrzehnt wird der Aufstieg der Motown-Musik assoziiert?', 'es': '¿Con que decada se asocia el auge de la musica Motown?', 'nl': 'Met welk decennium wordt de opkomst van Motownmuziek geassocieerd?', 'sv': 'Vilket artionde forknippas med framvaxten av Motown-musik?', 'fr': "A quelle decennie associe-t-on la montee de la musique Motown ?", 'da': 'Hvilket årti forbindes med fremkomsten af Motown-musik?'},
 'opt1': {'en': '1960s', 'de': '1960er', 'es': 'los años 60', 'nl': 'jaren 60', 'sv': '1960-talet', 'fr': 'les annees 1960', 'da': '1960’erne'},
 'opt2': {'en': '1950s', 'de': '1950er', 'es': 'los años 50', 'nl': 'jaren 50', 'sv': '1950-talet', 'fr': 'les annees 1950', 'da': '1950’erne'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Kyoto?', 'de': 'In welchem Land liegt die antike Stadt Kyoto?', 'es': '¿En que pais esta la antigua ciudad de Kioto?', 'nl': 'In welk land ligt de oude stad Kyoto?', 'sv': 'I vilket land ligger den antika staden Kyoto?', 'fr': "Dans quel pays se trouve l'ancienne cite de Kyoto ?", 'da': 'Hvilket land huser den antikke by Kyoto?'},
 'opt1': {'en': 'Japan', 'de': 'Japan', 'es': 'Japon', 'nl': 'Japan', 'sv': 'Japan', 'fr': 'le Japon', 'da': 'Japan'},
 'opt2': {'en': 'China', 'de': 'China', 'es': 'China', 'nl': 'China', 'sv': 'Kina', 'fr': 'la Chine', 'da': 'Kina'}},

{'category': 'Animals', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which farm animal is known for mooing and giving milk?', 'de': 'Welches Nutztier muht und gibt Milch?', 'es': '¿Que animal de granja muge y da leche?', 'nl': 'Welk boerderijdier loeit en geeft melk?', 'sv': 'Vilket lantbruksdjur bolar och ger mjolk?', 'fr': "Quel animal de ferme meugle et donne du lait ?", 'da': 'Hvilket landbrugsdyr brøler og giver mælk?'},
 'opt1': {'en': 'Cow', 'de': 'Kuh', 'es': 'Vaca', 'nl': 'Koe', 'sv': 'Ko', 'fr': 'Vache', 'da': 'Ko'},
 'opt2': {'en': 'Buffalo', 'de': 'Buffel', 'es': 'Bufalo', 'nl': 'Buffel', 'sv': 'Buffel', 'fr': 'Buffle', 'da': 'Bøffel'}},

{'category': 'Nature', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which gemstone is the birthstone for those born in March?', 'de': 'Welcher Edelstein ist der Geburtsstein fuer im Maerz Geborene?', 'es': '¿Que piedra preciosa es la piedra de nacimiento para los nacidos en marzo?', 'nl': 'Welke edelsteen is de geboortesteen voor mensen geboren in maart?', 'sv': 'Vilken adelsten ar fodelsestenen for de fodda i mars?', 'fr': "Quelle pierre precieuse est la pierre de naissance de mars ?", 'da': 'Hvilken ædelsten er fødselsstenen for dem, der er født i marts?'},
 'opt1': {'en': 'Aquamarine', 'de': 'Aquamarin', 'es': 'Aguamarina', 'nl': 'Aquamarijn', 'sv': 'Akvamarin', 'fr': 'Aigue-marine', 'da': 'Akvamarin'},
 'opt2': {'en': 'Amethyst', 'de': 'Amethyst', 'es': 'Amatista', 'nl': 'Amethist', 'sv': 'Ametist', 'fr': 'Amethyste', 'da': 'Ametyst'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which classic outdoor activity involves flying a fabric or plastic object on a string in the wind?', 'de': 'Bei welcher klassischen Freiluftaktivitaet laesst man ein Stoff- oder Plastikobjekt an einer Schnur im Wind fliegen?', 'es': '¿Que actividad clasica al aire libre implica hacer volar un objeto de tela o plastico atado a un hilo en el viento?', 'nl': 'Bij welke klassieke buitenactiviteit laat je een stoffen of plastic object aan een touwtje in de wind vliegen?', 'sv': 'Vilken klassisk utomhusaktivitet innebar att flyga ett tyg- eller plastforemal pa ett snore i vinden?', 'fr': "Quelle activite de plein air classique consiste a faire voler un objet en tissu ou en plastique attache a une ficelle dans le vent ?", 'da': 'Hvilken klassisk udendørsaktivitet involverer at flyve en genstand af stof eller plastik på en snor i vinden?'},
 'opt1': {'en': 'Frisbee', 'de': 'Frisbee', 'es': 'Frisbee', 'nl': 'Frisbee', 'sv': 'Frisbee', 'fr': 'Frisbee', 'da': 'Frisbee'},
 'opt2': {'en': 'Kite flying', 'de': 'Drachensteigen', 'es': 'Volar cometas', 'nl': 'Vliegeren', 'sv': 'Draken kastning', 'fr': "Le vol de cerf-volant", 'da': 'Drageflyvning'}},

{'category': 'Food', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which green herb is used to flavor traditional Thanksgiving stuffing?', 'de': 'Welches gruene Kraut wird verwendet, um traditionelle Thanksgiving-Fuellung zu wuerzen?', 'es': '¿Que hierba verde se usa para sazonar el relleno tradicional de Accion de Gracias?', 'nl': 'Welk groen kruid wordt gebruikt om traditionele Thanksgiving-vulling te kruiden?', 'sv': 'Vilken gron ort anvands for att smaksatta traditionell Thanksgiving-fyllning?', 'fr': "Quelle herbe verte est utilisee pour aromatiser la farce traditionnelle de Thanksgiving ?", 'da': 'Hvilken grøn urt bruges til at smage traditionel Thanksgiving-fyld?'},
 'opt1': {'en': 'Thyme', 'de': 'Thymian', 'es': 'Tomillo', 'nl': 'Tijm', 'sv': 'Timjan', 'fr': 'Le thym', 'da': 'Timian'},
 'opt2': {'en': 'Sage', 'de': 'Salbei', 'es': 'Salvia', 'nl': 'Salie', 'sv': 'Salvia', 'fr': 'La sauge', 'da': 'Salvie'}},

{'category': 'Culture', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which flower is traditionally associated with the birth month of January?', 'de': 'Welche Blume wird traditionell mit dem Geburtsmonat Januar assoziiert?', 'es': '¿Que flor se asocia tradicionalmente con el mes de nacimiento de enero?', 'nl': 'Welke bloem wordt traditioneel geassocieerd met de geboortemaand januari?', 'sv': 'Vilken blomma forknippas traditionellt med fodelsemanaden januari?', 'fr': "Quelle fleur est traditionnellement associee au mois de naissance de janvier ?", 'da': 'Hvilken blomst forbindes traditionelt med fødselsmåneden januar?'},
 'opt1': {'en': 'Snowdrop', 'de': 'Schneegloeckchen', 'es': 'Campanilla de invierno', 'nl': 'Sneeuwklokje', 'sv': 'Snodroppe', 'fr': 'Perce-neige', 'da': 'Vintergæk'},
 'opt2': {'en': 'Carnation', 'de': 'Nelke', 'es': 'Clavel', 'nl': 'Anjer', 'sv': 'Nejlika', 'fr': 'Oeillet', 'da': 'Nellike'}},

{'category': 'Geography', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which U.S. state is home to the Everglades National Park?', 'de': 'In welchem US-Bundesstaat liegt der Everglades-Nationalpark?', 'es': '¿En que estado de EE.UU. esta el Parque Nacional Everglades?', 'nl': 'In welke Amerikaanse staat ligt het Everglades National Park?', 'sv': 'I vilken amerikansk delstat ligger Everglades nationalpark?', 'fr': "Dans quel État americain se trouve le parc national des Everglades ?", 'da': 'I hvilken amerikansk delstat ligger Everglades nationalpark?'},
 'opt1': {'en': 'Georgia', 'de': 'Georgia', 'es': 'Georgia', 'nl': 'Georgia', 'sv': 'Georgia', 'fr': 'la Georgie', 'da': 'Georgia'},
 'opt2': {'en': 'Florida', 'de': 'Florida', 'es': 'Florida', 'nl': 'Florida', 'sv': 'Florida', 'fr': 'la Floride', 'da': 'Florida'}},

{'category': 'History', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which craft involves creating jewelry by stringing beads onto thread or wire?', 'de': 'Bei welchem Handwerk wird Schmuck durch das Aufreihen von Perlen auf Faden oder Draht hergestellt?', 'es': '¿Que manualidad implica crear joyeria enhebrando cuentas en hilo o alambre?', 'nl': 'Bij welk handwerk maak je sieraden door kralen aan een draad of draadje te rijgen?', 'sv': 'Vilket hantverk innebar att skapa smycken genom att trada palror pa trad eller strang?', 'fr': "Quel art manuel consiste a creer des bijoux en enfilant des perles sur un fil ou un cable ?", 'da': 'Hvilket håndværk involverer at skabe smykker ved at tråde perler på tråd eller wire?'},
 'opt1': {'en': 'Wire wrapping', 'de': 'Drahtwickeltechnik', 'es': 'Enrollado de alambre', 'nl': 'Draadwikkelen', 'sv': 'Tradlindning', 'fr': 'Le wire wrapping', 'da': 'Trådvikling'},
 'opt2': {'en': 'Beading', 'de': 'Perlenarbeit', 'es': 'Abalorios', 'nl': 'Kralenwerk', 'sv': 'Palrarbete', 'fr': 'Le perlage', 'da': 'Perlearbejde'}},

{'category': 'Sports', 'tier': 1, 'correct_index': 2,
 'question': {'en': 'Which sport involves throwing a javelin as far as possible?', 'de': 'Bei welcher Sportart wirft man einen Speer so weit wie moeglich?', 'es': '¿Que deporte consiste en lanzar una jabalina lo mas lejos posible?', 'nl': 'Bij welke sport gooi je een speer zo ver mogelijk?', 'sv': 'Vilken sport innebar att kasta ett spjut sa langt som mojligt?', 'fr': "Quel sport consiste a lancer un javelot le plus loin possible ?", 'da': 'Hvilken sport involverer at kaste et spyd så langt som muligt?'},
 'opt1': {'en': 'Shot put', 'de': 'Kugelstossen', 'es': 'Lanzamiento de peso', 'nl': 'Kogelstoten', 'sv': 'Kulstotning', 'fr': 'Le lancer de poids', 'da': 'Kuglestød'},
 'opt2': {'en': 'Javelin throw', 'de': 'Speerwurf', 'es': 'Lanzamiento de jabalina', 'nl': 'Speerwerpen', 'sv': 'Spjutkastning', 'fr': 'Le lancer de javelot', 'da': 'Spydkast'}},

{'category': 'Fashion', 'tier': 1, 'correct_index': 1,
 'question': {'en': 'Which item of clothing is worn on the feet in the shower or at the beach?', 'de': 'Welches Kleidungsstueck wird an den Fuessen unter der Dusche oder am Strand getragen?', 'es': '¿Que prenda se lleva en los pies en la ducha o en la playa?', 'nl': 'Welk kledingstuk wordt aan de voeten gedragen onder de douche of op het strand?', 'sv': 'Vilket klädesplagg baers pa fotterna i duschen eller pa stranden?', 'fr': "Quel vetement se porte aux pieds sous la douche ou a la plage ?", 'da': 'Hvilken beklædningsgenstand bæres på fødderne i bruseren eller på stranden?'},
 'opt1': {'en': 'Sandals', 'de': 'Sandalen', 'es': 'Sandalias', 'nl': 'Sandalen', 'sv': 'Sandaler', 'fr': 'Sandales', 'da': 'Sandaler'},
 'opt2': {'en': 'Flip-flops', 'de': 'Flip-Flops', 'es': 'Chanclas', 'nl': 'Slippers', 'sv': 'Flip-flops', 'fr': 'Tongs', 'da': 'Flip-flops'}},

{'category': 'Entertainment', 'tier': 1, 'correct_index': 1,
 'question': {'en': "Which classic cartoon features a rabbit who often says 'What's up, Doc?'", 'de': "In welchem klassischen Cartoon kommt ein Kaninchen vor, das oft 'Was geht ab, Doc?' sagt?", 'es': "¿Que dibujo animado clasico presenta a un conejo que a menudo dice '¿Que hay de nuevo, viejo?'", 'nl': "In welke klassieke tekenfilm komt een konijn voor dat vaak zegt 'Wat is er, Doc?'", 'sv': "I vilken klassisk tecknad film forekommer en kanin som ofta sager 'Vad star pa, doc?'", 'fr': "Quel dessin anime classique met en scene un lapin qui dit souvent 'Quoi de neuf, docteur ?'", 'da': "Hvilken klassisk tegnefilm har en kanin, der ofte siger 'Hvad så, Doc?'"},
 'opt1': {'en': 'Bugs Bunny', 'de': 'Bugs Bunny', 'es': 'Bugs Bunny', 'nl': 'Bugs Bunny', 'sv': 'Bugs Bunny', 'fr': 'Bugs Bunny', 'da': 'Snurre Snup'},
 'opt2': {'en': 'Roger Rabbit', 'de': 'Roger Rabbit', 'es': 'Roger Rabbit', 'nl': 'Roger Rabbit', 'sv': 'Roger Rabbit', 'fr': 'Roger Rabbit', 'da': 'Roger Rabbit'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Hue, a former imperial capital?', 'de': 'In welchem Land liegt die antike Stadt Hue, eine ehemalige kaiserliche Hauptstadt?', 'es': '¿En que pais esta la antigua ciudad de Hue, una antigua capital imperial?', 'nl': 'In welk land ligt de oude stad Hue, een voormalige keizerlijke hoofdstad?', 'sv': 'I vilket land ligger den antika staden Hue, en tidigare kejserlig huvudstad?', 'fr': "Dans quel pays se trouve l'ancienne cite de Hue, ancienne capitale imperiale ?", 'da': 'Hvilket land huser den antikke by Hue, en tidligere kejserlig hovedstad?'},
 'opt1': {'en': 'Cambodia', 'de': 'Kambodscha', 'es': 'Camboya', 'nl': 'Cambodja', 'sv': 'Kambodja', 'fr': 'le Cambodge', 'da': 'Cambodja'},
 'opt2': {'en': 'Vietnam', 'de': 'Vietnam', 'es': 'Vietnam', 'nl': 'Vietnam', 'sv': 'Vietnam', 'fr': 'le Vietnam', 'da': 'Vietnam'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which First Lady was known for redesigning the White House Rose Garden?', 'de': 'Welche First Lady war bekannt fuer die Neugestaltung des Rosengartens des Weissen Hauses?', 'es': '¿Que primera dama fue conocida por rediseñar el Jardin de Rosas de la Casa Blanca?', 'nl': 'Welke First Lady stond bekend om het herontwerpen van de Rozentuin van het Witte Huis?', 'sv': 'Vilken First Lady var kand for att designa om Vita husets rosentradgard?', 'fr': "Quelle Première dame etait connue pour avoir redessine le jardin des roses de la Maison-Blanche ?", 'da': 'Hvilken førstedame var kendt for at redesigne Det Hvide Hus’ rosenhave?'},
 'opt1': {'en': 'Pat Nixon', 'de': 'Pat Nixon', 'es': 'Pat Nixon', 'nl': 'Pat Nixon', 'sv': 'Pat Nixon', 'fr': 'Pat Nixon', 'da': 'Pat Nixon'},
 'opt2': {'en': 'Jacqueline Kennedy', 'de': 'Jacqueline Kennedy', 'es': 'Jacqueline Kennedy', 'nl': 'Jacqueline Kennedy', 'sv': 'Jacqueline Kennedy', 'fr': 'Jacqueline Kennedy', 'da': 'Jacqueline Kennedy'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Malaysian dish is a spicy noodle soup with coconut milk and shrimp?', 'de': 'Welches malaysische Gericht ist eine scharfe Nudelsuppe mit Kokosmilch und Garnelen?', 'es': '¿Que plato malayo es una sopa de fideos picante con leche de coco y camarones?', 'nl': 'Welk Maleisisch gerecht is een pittige noedelsoep met kokosmelk en garnalen?', 'sv': 'Vilken malaysisk ratt ar en kryddig nudelsoppa med kokosmjolk och raga?', 'fr': "Quel plat malaisien est une soupe de nouilles epicee au lait de coco et aux crevettes ?", 'da': 'Hvilken malaysisk ret er en krydret nudelsuppe med kokosmælk og rejer?'},
 'opt1': {'en': 'Mee goreng', 'de': 'Mee Goreng', 'es': 'Mee goreng', 'nl': 'Mee goreng', 'sv': 'Mee goreng', 'fr': 'Mee goreng', 'da': 'Mee goreng'},
 'opt2': {'en': 'Laksa', 'de': 'Laksa', 'es': 'Laksa', 'nl': 'Laksa', 'sv': 'Laksa', 'fr': 'Laksa', 'da': 'Laksa'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which body system produces hormones, including the thyroid and adrenal glands?', 'de': 'Welches Koerpersystem produziert Hormone, einschliesslich der Schilddruese und Nebenniere?', 'es': '¿Que sistema corporal produce hormonas, incluidas las glandulas tiroides y suprarrenales?', 'nl': 'Welk lichaamssysteem produceert hormonen, waaronder de schildklier en bijnieren?', 'sv': 'Vilket kroppssystem producerar hormoner, inklusive skoldkorteln och binjurarna?', 'fr': "Quel systeme corporel produit des hormones, y compris la thyroide et les glandes surrenales ?", 'da': 'Hvilket kropssystem producerer hormoner, herunder skjoldbruskkirtlen og binyrerne?'},
 'opt1': {'en': 'Nervous system', 'de': 'Nervensystem', 'es': 'Sistema nervioso', 'nl': 'Zenuwstelsel', 'sv': 'Nervsystemet', 'fr': 'Le systeme nerveux', 'da': 'Nervesystemet'},
 'opt2': {'en': 'Endocrine system', 'de': 'Endokrines System', 'es': 'Sistema endocrino', 'nl': 'Hormoonstelsel', 'sv': 'Endokrina systemet', 'fr': 'Le systeme endocrinien', 'da': 'Det endokrine system'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'Little House on the Prairie'?", 'de': "Welche Autorin schrieb 'Unsere kleine Farm'?", 'es': "¿Que autora escribio 'La casa de la pradera'?", 'nl': "Welke auteur schreef 'Het Kleine Huis in het Grote Bos'?", 'sv': "Vilken forfattare skrev 'Lilla huset pa prarien'?", 'fr': "Quelle autrice a ecrit 'La Petite Maison dans la prairie' ?", 'da': "Hvilken forfatter skrev 'Det lille hus på prærien'?"},
 'opt1': {'en': 'Beverly Cleary', 'de': 'Beverly Cleary', 'es': 'Beverly Cleary', 'nl': 'Beverly Cleary', 'sv': 'Beverly Cleary', 'fr': 'Beverly Cleary', 'da': 'Beverly Cleary'},
 'opt2': {'en': 'Laura Ingalls Wilder', 'de': 'Laura Ingalls Wilder', 'es': 'Laura Ingalls Wilder', 'nl': 'Laura Ingalls Wilder', 'sv': 'Laura Ingalls Wilder', 'fr': 'Laura Ingalls Wilder', 'da': 'Laura Ingalls Wilder'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which actress played Regina George in the film 'Mean Girls'?", 'de': "Welche Schauspielerin spielte Regina George im Film 'Mean Girls'?", 'es': "¿Que actriz interpreto a Regina George en la pelicula 'Chicas malas'?", 'nl': "Welke actrice speelde Regina George in de film 'Mean Girls'?", 'sv': "Vilken skadespelerska spelade Regina George i filmen 'Mean Girls'?", 'fr': "Quelle actrice a joue Regina George dans le film 'Lolita malgre moi' ?", 'da': "Hvilken skuespiller spillede Regina George i filmen 'Mean Girls'?"},
 'opt1': {'en': 'Amanda Seyfried', 'de': 'Amanda Seyfried', 'es': 'Amanda Seyfried', 'nl': 'Amanda Seyfried', 'sv': 'Amanda Seyfried', 'fr': 'Amanda Seyfried', 'da': 'Amanda Seyfried'},
 'opt2': {'en': 'Rachel McAdams', 'de': 'Rachel McAdams', 'es': 'Rachel McAdams', 'nl': 'Rachel McAdams', 'sv': 'Rachel McAdams', 'fr': 'Rachel McAdams', 'da': 'Rachel McAdams'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which singer is known as the 'First Lady of Song' for her jazz vocals?", 'de': "Welche Saengerin wird fuer ihren Jazzgesang 'First Lady of Song' genannt?", 'es': "¿Que cantante es conocida como la 'Primera Dama de la Canción' por su voz de jazz?", 'nl': "Welke zangeres staat bekend als de 'First Lady of Song' vanwege haar jazzstem?", 'sv': "Vilken sangerska kallas 'First Lady of Song' for sin jazzsang?", 'fr': "Quelle chanteuse est connue comme la 'Première dame du chant' pour sa voix jazz ?", 'da': "Hvilken sanger er kendt som 'First Lady of Song' for sin jazzstemme?"},
 'opt1': {'en': 'Billie Holiday', 'de': 'Billie Holiday', 'es': 'Billie Holiday', 'nl': 'Billie Holiday', 'sv': 'Billie Holiday', 'fr': 'Billie Holiday', 'da': 'Billie Holiday'},
 'opt2': {'en': 'Ella Fitzgerald', 'de': 'Ella Fitzgerald', 'es': 'Ella Fitzgerald', 'nl': 'Ella Fitzgerald', 'sv': 'Ella Fitzgerald', 'fr': 'Ella Fitzgerald', 'da': 'Ella Fitzgerald'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Hoi An, known for lanterns?', 'de': 'In welchem Land liegt die antike Stadt Hoi An, bekannt fuer ihre Laternen?', 'es': '¿En que pais esta la antigua ciudad de Hoi An, conocida por sus faroles?', 'nl': 'In welk land ligt de oude stad Hoi An, bekend om zijn lantaarns?', 'sv': 'I vilket land ligger den antika staden Hoi An, kand for sina lyktor?', 'fr': "Dans quel pays se trouve l'ancienne cite de Hoi An, connue pour ses lanternes ?", 'da': 'Hvilket land huser den antikke by Hoi An, kendt for sine lanterner?'},
 'opt1': {'en': 'Vietnam', 'de': 'Vietnam', 'es': 'Vietnam', 'nl': 'Vietnam', 'sv': 'Vietnam', 'fr': 'le Vietnam', 'da': 'Vietnam'},
 'opt2': {'en': 'Thailand', 'de': 'Thailand', 'es': 'Tailandia', 'nl': 'Thailand', 'sv': 'Thailand', 'fr': 'la Thailande', 'da': 'Thailand'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large mammal has a hump for storing fat and lives in deserts?', 'de': 'Welches grosse Saeugetier hat einen Hoecker zum Speichern von Fett und lebt in Wuesten?', 'es': '¿Que gran mamifero tiene una joroba para almacenar grasa y vive en desiertos?', 'nl': 'Welk groot zoogdier heeft een bult om vet op te slaan en leeft in woestijnen?', 'sv': 'Vilket stort daggdjur har en puckel for att lagra fett och lever i okenar?', 'fr': "Quel grand mammifere a une bosse pour stocker la graisse et vit dans les deserts ?", 'da': 'Hvilket stort pattedyr har en pukkel til at gemme fedt og lever i ørkener?'},
 'opt1': {'en': 'Camel', 'de': 'Kamel', 'es': 'Camello', 'nl': 'Kameel', 'sv': 'Kamel', 'fr': 'Chameau', 'da': 'Kamel'},
 'opt2': {'en': 'Llama', 'de': 'Lama', 'es': 'Llama', 'nl': 'Lama', 'sv': 'Lama', 'fr': 'Lama', 'da': 'Lama'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Filipino dessert is shaved ice with sweet beans, fruit, and evaporated milk?', 'de': 'Welches philippinische Dessert ist geschabtes Eis mit suessen Bohnen, Obst und Kondensmilch?', 'es': '¿Que postre filipino es hielo raspado con frijoles dulces, fruta y leche evaporada?', 'nl': 'Welk Filipijns dessert is geschraapt ijs met zoete bonen, fruit en gecondenseerde melk?', 'sv': 'Vilken filippinsk dessert ar rivet is med sota bonor, frukt och kondenserad mjolk?', 'fr': "Quel dessert philippin est de la glace pilee avec des haricots sucres, des fruits et du lait concentre ?", 'da': 'Hvilken filippinsk dessert er skrabet is med søde bønner, frugt og kondenseret mælk?'},
 'opt1': {'en': 'Halo-halo', 'de': 'Halo-Halo', 'es': 'Halo-halo', 'nl': 'Halo-halo', 'sv': 'Halo-halo', 'fr': 'Halo-halo', 'da': 'Halo-halo'},
 'opt2': {'en': 'Turon', 'de': 'Turon', 'es': 'Turon', 'nl': 'Turon', 'sv': 'Turon', 'fr': 'Turon', 'da': 'Turon'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which First Lady was known for her extensive work with Vietnam War veterans and POW families?', 'de': 'Welche First Lady war bekannt fuer ihre umfangreiche Arbeit mit Vietnamkrieg-Veteranen und Familien von Kriegsgefangenen?', 'es': '¿Que primera dama fue conocida por su extenso trabajo con veteranos de la Guerra de Vietnam y familias de prisioneros de guerra?', 'nl': 'Welke First Lady stond bekend om haar uitgebreide werk met veteranen van de Vietnamoorlog en krijgsgevangenenfamilies?', 'sv': 'Vilken First Lady var kand for sitt omfattande arbete med veteraner fran Vietnamkriget och krigsfangars familjer?', 'fr': "Quelle Première dame etait connue pour son important travail avec les veterans de la guerre du Vietnam et les familles de prisonniers de guerre ?", 'da': 'Hvilken førstedame var kendt for sit omfattende arbejde med veteraner fra Vietnamkrigen og krigsfangefamilier?'},
 'opt1': {'en': 'Betty Ford', 'de': 'Betty Ford', 'es': 'Betty Ford', 'nl': 'Betty Ford', 'sv': 'Betty Ford', 'fr': 'Betty Ford', 'da': 'Betty Ford'},
 'opt2': {'en': 'Pat Nixon', 'de': 'Pat Nixon', 'es': 'Pat Nixon', 'nl': 'Pat Nixon', 'sv': 'Pat Nixon', 'fr': 'Pat Nixon', 'da': 'Pat Nixon'}},

{'category': 'Arts', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which artist is known for the geometric abstract style using primary colors and black grid lines?', 'de': 'Welcher Kuenstler ist fuer den geometrisch abstrakten Stil mit Primaerfarben und schwarzen Gitterlinien bekannt?', 'es': '¿Que artista es conocido por el estilo abstracto geometrico usando colores primarios y lineas de rejilla negras?', 'nl': 'Welke kunstenaar staat bekend om de geometrische abstracte stijl met primaire kleuren en zwarte rasterlijnen?', 'sv': 'Vilken konstnar ar kand for den geometriska abstrakta stilen med primarfarger och svarta rutlinjer?', 'fr': "Quel artiste est connu pour son style abstrait geometrique utilisant les couleurs primaires et des lignes de grille noires ?", 'da': 'Hvilken kunstner er kendt for den geometriske abstrakte stil med primærfarver og sorte gitterlinjer?'},
 'opt1': {'en': 'Piet Mondrian', 'de': 'Piet Mondrian', 'es': 'Piet Mondrian', 'nl': 'Piet Mondriaan', 'sv': 'Piet Mondrian', 'fr': 'Piet Mondrian', 'da': 'Piet Mondrian'},
 'opt2': {'en': 'Wassily Kandinsky', 'de': 'Wassily Kandinsky', 'es': 'Wassily Kandinsky', 'nl': 'Wassily Kandinsky', 'sv': 'Wassily Kandinsky', 'fr': 'Wassily Kandinsky', 'da': 'Wassily Kandinsky'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which flowering plant's blooms open and close with the sun, often called a 'sun worshipper'?", 'de': "Welche bluehende Pflanze oeffnet und schliesst ihre Blueten mit der Sonne, oft 'Sonnenanbeter' genannt?", 'es': "¿Que planta con flores abre y cierra sus flores con el sol, a menudo llamada 'adoradora del sol'?", 'nl': "Welke bloeiende plant opent en sluit zijn bloesems met de zon, vaak 'zonaanbidder' genoemd?", 'sv': "Vilken blommande vaxts blommor oppnar och stanger sig med solen, ofta kallad 'soldyrkare'?", 'fr': "Quelle plante a fleurs ouvre et ferme ses fleurs avec le soleil, souvent appelee 'adoratrice du soleil' ?", 'da': "Hvilken blomstrende plantes blomster åbner og lukker med solen, ofte kaldet 'soltilbeder'?"},
 'opt1': {'en': 'Marigold', 'de': 'Ringelblume', 'es': 'Caléndula', 'nl': 'Goudsbloem', 'sv': 'Ringblomma', 'fr': 'Souci', 'da': 'Morgenfrue'},
 'opt2': {'en': 'Sunflower', 'de': 'Sonnenblume', 'es': 'Girasol', 'nl': 'Zonnebloem', 'sv': 'Solros', 'fr': 'Tournesol', 'da': 'Solsikke'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which musical features the song 'Suddenly Seymour' about a carnivorous plant?", 'de': "In welchem Musical kommt das Lied 'Suddenly Seymour' ueber eine fleischfressende Pflanze vor?", 'es': "¿Que musical presenta la cancion 'Suddenly Seymour' sobre una planta carnivora?", 'nl': "Welke musical bevat het lied 'Suddenly Seymour' over een vleesetende plant?", 'sv': "Vilken musikal innehaller latten 'Suddenly Seymour' om en kottätande vaxt?", 'fr': "Quelle comedie musicale comporte la chanson 'Suddenly Seymour' au sujet d'une plante carnivore ?", 'da': "Hvilken musical har sangen 'Suddenly Seymour' om en kødædende plante?"},
 'opt1': {'en': 'Little Shop of Horrors', 'de': 'Der kleine Horrorladen', 'es': 'La pequeña tienda de los horrores', 'nl': 'Little Shop of Horrors', 'sv': 'Lilla affaren av skrack', 'fr': 'La Petite Boutique des horreurs', 'da': 'Den lille rædselsbutik'},
 'opt2': {'en': 'Sweeney Todd', 'de': 'Sweeney Todd', 'es': 'Sweeney Todd', 'nl': 'Sweeney Todd', 'sv': 'Sweeney Todd', 'fr': 'Sweeney Todd', 'da': 'Sweeney Todd'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country has Stockholm as its capital?', 'de': 'Welches Land hat Stockholm als Hauptstadt?', 'es': '¿Que pais tiene a Estocolmo como su capital?', 'nl': 'Welk land heeft Stockholm als hoofdstad?', 'sv': 'Vilket land har Stockholm som huvudstad?', 'fr': "Quel pays a Stockholm pour capitale ?", 'da': 'Hvilket land har Stockholm som hovedstad?'},
 'opt1': {'en': 'Norway', 'de': 'Norwegen', 'es': 'Noruega', 'nl': 'Noorwegen', 'sv': 'Norge', 'fr': 'la Norvège', 'da': 'Norge'},
 'opt2': {'en': 'Sweden', 'de': 'Schweden', 'es': 'Suecia', 'nl': 'Zweden', 'sv': 'Sverige', 'fr': 'la Suede', 'da': 'Sverige'}},

{'category': 'Food', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which spice is used in traditional Vietnamese pho broth for its aromatic flavor?', 'de': 'Welches Gewuerz wird in traditioneller vietnamesischer Pho-Bruehe fuer ihren aromatischen Geschmack verwendet?', 'es': '¿Que especia se usa en el caldo de pho vietnamita tradicional por su sabor aromatico?', 'nl': 'Welke specerij wordt gebruikt in traditionele Vietnamese pho-bouillon voor zijn aromatische smaak?', 'sv': 'Vilken krydda anvands i traditionell vietnamesisk pho-buljong for sin aromatiska smak?', 'fr': "Quelle epice est utilisee dans le bouillon de pho vietnamien traditionnel pour sa saveur aromatique ?", 'da': 'Hvilket krydderi bruges i traditionel vietnamesisk pho-bouillon for sin aromatiske smag?'},
 'opt1': {'en': 'Cardamom', 'de': 'Kardamom', 'es': 'Cardamomo', 'nl': 'Kardemom', 'sv': 'Kardemumma', 'fr': 'La cardamome', 'da': 'Kardemomme'},
 'opt2': {'en': 'Star anise', 'de': 'Sternanis', 'es': 'Anis estrellado', 'nl': 'Steranijs', 'sv': 'Stjaranis', 'fr': 'Anis etoile', 'da': 'Stjerneanis'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which craft involves creating pictures using yarn glued onto a board in patterns, popular in Mexican folk art?', 'de': 'Bei welchem Handwerk werden Bilder aus Garn erstellt, das in Mustern auf ein Brett geklebt wird, beliebt in mexikanischer Volkskunst?', 'es': '¿Que manualidad implica crear imagenes con hilo pegado en patrones sobre una tabla, popular en el arte popular mexicano?', 'nl': 'Bij welk handwerk creëer je afbeeldingen met garen dat in patronen op een bord wordt gelijmd, populair in Mexicaanse volkskunst?', 'sv': 'Vilket hantverk innebar att skapa bilder med garn limmat i monster pa en skiva, populart i mexikansk folkkonst?', 'fr': "Quel art manuel consiste a creer des images avec du fil colle en motifs sur une planche, populaire dans l'art populaire mexicain ?", 'da': 'Hvilket håndværk involverer at skabe billeder med garn limet i mønstre på et bræt, populært i mexicansk folkekunst?'},
 'opt1': {'en': 'Yarn painting', 'de': 'Garnmalerei', 'es': 'Cuadro de estambre', 'nl': 'Garenschilderen', 'sv': 'Garnmalning', 'fr': 'La peinture au fil', 'da': 'Garnmaleri'},
 'opt2': {'en': 'Weaving', 'de': 'Weben', 'es': 'Tejido en telar', 'nl': 'Weven', 'sv': 'Vavning', 'fr': 'Le tissage', 'da': 'Vævning'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which part of the digestive system absorbs most nutrients from food?', 'de': 'Welcher Teil des Verdauungssystems nimmt die meisten Naehrstoffe aus der Nahrung auf?', 'es': '¿Que parte del sistema digestivo absorbe la mayoria de los nutrientes de los alimentos?', 'nl': 'Welk deel van het spijsverteringsstelsel absorbeert de meeste voedingsstoffen uit voedsel?', 'sv': 'Vilken del av matsmaltningssystemet absorberar de flesta naringsamnen fran maten?', 'fr': "Quelle partie du systeme digestif absorbe la plupart des nutriments des aliments ?", 'da': 'Hvilken del af fordøjelsessystemet optager de fleste næringsstoffer fra maden?'},
 'opt1': {'en': 'Small intestine', 'de': 'Duenndarm', 'es': 'Intestino delgado', 'nl': 'Dunne darm', 'sv': 'Tunntarmen', 'fr': "L'intestin grele", 'da': 'Tyndtarm'},
 'opt2': {'en': 'Large intestine', 'de': 'Dickdarm', 'es': 'Intestino grueso', 'nl': 'Dikke darm', 'sv': 'Tjocktarmen', 'fr': 'Le gros intestin', 'da': 'Tyktarm'}},

{'category': 'Sports', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'In diving, what is the term for a dive with three and a half somersaults?', 'de': 'Wie nennt man beim Wasserspringen einen Sprung mit drei und einer halben Rolle?', 'es': '¿En clavados, como se llama un salto con tres saltos mortales y medio?', 'nl': 'Hoe heet bij schoonspringen een duik met drieënhalve salto?', 'sv': 'Vad kallas i simhopp ett hopp med tre och en halv volter?', 'fr': "En plongeon, comment appelle-t-on un plongeon avec trois saltos et demi ?", 'da': 'Hvad kaldes i udspring et spring med tre og en halv saltomortale?'},
 'opt1': {'en': 'Forward pike', 'de': 'Vorwaerts-Gehechtet', 'es': 'Adelante en carpa', 'nl': 'Voorwaarts gehoekt', 'sv': 'Framatvant vinkelposition', 'fr': 'Avant carpe', 'da': 'Fremad-piket'},
 'opt2': {'en': 'Reverse three and a half', 'de': 'Reverse-Dreieinhalbfacher', 'es': 'Tres y medio invertido', 'nl': 'Achterwaartse driëenhalve', 'sv': 'Bakvant tre och en halv', 'fr': "Trois et demi en arriere", 'da': 'Omvendt tre-og-en-halv'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which Chinese city is known for its ancient Forbidden City?", 'de': 'Welche chinesische Stadt ist fuer ihre antike Verbotene Stadt bekannt?', 'es': '¿Que ciudad china es conocida por su antigua Ciudad Prohibida?', 'nl': 'Welke Chinese stad staat bekend om zijn oude Verboden Stad?', 'sv': 'Vilken kinesisk stad ar kand for sin antika Forbjudna staden?', 'fr': "Quelle ville chinoise est connue pour son ancienne Cite interdite ?", 'da': 'Hvilken kinesisk by er kendt for sin antikke Forbudte By?'},
 'opt1': {'en': 'Beijing', 'de': 'Peking', 'es': 'Pekin', 'nl': 'Peking', 'sv': 'Peking', 'fr': 'Pekin', 'da': 'Beijing'},
 'opt2': {'en': 'Shanghai', 'de': 'Shanghai', 'es': 'Shanghai', 'nl': 'Shanghai', 'sv': 'Shanghai', 'fr': 'Shanghai', 'da': 'Shanghai'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which actress played Diana Prince in the 'Wonder Woman' films?", 'de': "Welche Schauspielerin spielte Diana Prince in den 'Wonder Woman'-Filmen?", 'es': "¿Que actriz interpreto a Diana Prince en las peliculas de 'Wonder Woman'?", 'nl': "Welke actrice speelde Diana Prince in de 'Wonder Woman'-films?", 'sv': "Vilken skadespelerska spelade Diana Prince i 'Wonder Woman'-filmerna?", 'fr': "Quelle actrice a joue Diana Prince dans les films 'Wonder Woman' ?", 'da': "Hvilken skuespiller spillede Diana Prince i 'Wonder Woman'-filmene?"},
 'opt1': {'en': 'Gal Gadot', 'de': 'Gal Gadot', 'es': 'Gal Gadot', 'nl': 'Gal Gadot', 'sv': 'Gal Gadot', 'fr': 'Gal Gadot', 'da': 'Gal Gadot'},
 'opt2': {'en': 'Lynda Carter', 'de': 'Lynda Carter', 'es': 'Lynda Carter', 'nl': 'Lynda Carter', 'sv': 'Lynda Carter', 'fr': 'Lynda Carter', 'da': 'Lynda Carter'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which classic novel features the character Huckleberry Finn?', 'de': 'In welchem Klassiker kommt die Figur Huckleberry Finn vor?', 'es': '¿Que novela clasica presenta al personaje Huckleberry Finn?', 'nl': 'In welke klassieke roman komt het personage Huckleberry Finn voor?', 'sv': 'I vilken klassisk roman forekommer karaktaren Huckleberry Finn?', 'fr': "Quel roman classique met en scene le personnage de Huckleberry Finn ?", 'da': 'Hvilken klassisk roman har karakteren Huckleberry Finn?'},
 'opt1': {'en': 'The Adventures of Huckleberry Finn', 'de': 'Die Abenteuer des Huckleberry Finn', 'es': 'Las aventuras de Huckleberry Finn', 'nl': 'De Lotgevallen van Huckleberry Finn', 'sv': 'Huckleberry Finns aventyr', 'fr': 'Les Aventures de Huckleberry Finn', 'da': 'Huckleberry Finns oplevelser'},
 'opt2': {'en': 'Tom Sawyer Abroad', 'de': 'Tom Sawyer im Ausland', 'es': 'Tom Sawyer en el extranjero', 'nl': 'Tom Sawyer in het Buitenland', 'sv': 'Tom Sawyer utomlands', 'fr': "Tom Sawyer a l'etranger", 'da': 'Tom Sawyer i udlandet'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which 2000s singer is known for the song 'Since U Been Gone'?", 'de': "Welche Saengerin der 2000er ist fuer das Lied 'Since U Been Gone' bekannt?", 'es': "¿Que cantante de los años 2000 es conocida por la cancion 'Since U Been Gone'?", 'nl': "Welke zangeres uit de jaren 2000 staat bekend om het lied 'Since U Been Gone'?", 'sv': "Vilken sangerska fran 2000-talet ar kand for latten 'Since U Been Gone'?", 'fr': "Quelle chanteuse des annees 2000 est connue pour la chanson 'Since U Been Gone' ?", 'da': "Hvilken 2000’er-sanger er kendt for sangen 'Since U Been Gone'?"},
 'opt1': {'en': 'Kelly Clarkson', 'de': 'Kelly Clarkson', 'es': 'Kelly Clarkson', 'nl': 'Kelly Clarkson', 'sv': 'Kelly Clarkson', 'fr': 'Kelly Clarkson', 'da': 'Kelly Clarkson'},
 'opt2': {'en': 'Christina Aguilera', 'de': 'Christina Aguilera', 'es': 'Christina Aguilera', 'nl': 'Christina Aguilera', 'sv': 'Christina Aguilera', 'fr': 'Christina Aguilera', 'da': 'Christina Aguilera'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a rich casserole made with broccoli, cheese, and rice?', 'de': 'Welches suddstaatliche Gericht ist ein reichhaltiger Auflauf aus Brokkoli, Kaese und Reis?', 'es': '¿Que plato sureño de EE.UU. es una cacerola rica hecha con brocoli, queso y arroz?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een rijke ovenschotel gemaakt met broccoli, kaas en rijst?', 'sv': 'Vilken sydstatlig ratt ar en rik gryta gjord med broccoli, ost och ris?', 'fr': "Quel plat du Sud des États-Unis est un gratin riche fait avec du brocoli, du fromage et du riz ?", 'da': 'Hvilken sydstatsret er en rig gryderet lavet med broccoli, ost og ris?'},
 'opt1': {'en': 'Broccoli rice casserole', 'de': 'Brokkoli-Reis-Auflauf', 'es': 'Cacerola de brocoli y arroz', 'nl': 'Broccoli-rijstschotel', 'sv': 'Broccoli-risgryta', 'fr': 'Gratin de brocoli et riz', 'da': 'Broccoli-ris-gryderet'},
 'opt2': {'en': 'Green bean casserole', 'de': 'Gruene-Bohnen-Auflauf', 'es': 'Cacerola de judias verdes', 'nl': 'Sperziebonenschotel', 'sv': 'Grona bonor-gryta', 'fr': 'Le gratin de haricots verts', 'da': 'Grøn bønne-gryderet'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which queen ruled Portugal's colonial empire and was known as 'the Mad'?", 'de': "Welche Koenigin regierte Portugals Kolonialreich und wurde 'die Wahnsinnige' genannt?", 'es': "¿Que reina goberno el imperio colonial de Portugal y fue conocida como 'la Loca'?", 'nl': "Welke koningin regeerde over Portugals koloniale rijk en stond bekend als 'de Waanzinnige'?", 'sv': "Vilken drottning styrde Portugals kolonialrike och kallades 'den galna'?", 'fr': "Quelle reine a regne sur l'empire colonial portugais et etait connue comme 'la Folle' ?", 'da': "Hvilken dronning regerede Portugals kolonirige og var kendt som 'den gale'?"},
 'opt1': {'en': 'Catherine of Braganza', 'de': 'Katharina von Braganza', 'es': 'Catalina de Braganza', 'nl': 'Catharina van Braganza', 'sv': 'Katarina av Braganza', 'fr': 'Catherine de Bragance', 'da': 'Katarina af Braganza'},
 'opt2': {'en': 'Maria I of Portugal', 'de': 'Maria I. von Portugal', 'es': 'Maria I de Portugal', 'nl': 'Maria I van Portugal', 'sv': 'Maria I av Portugal', 'fr': 'Marie Ire de Portugal', 'da': 'Maria 1. af Portugal'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which large flightless bird is native to South America, related to the ostrich?', 'de': 'Welcher grosse flugunfaehige Vogel ist in Suedamerika heimisch und mit dem Strauss verwandt?', 'es': '¿Que gran ave no voladora es nativa de Sudamerica, emparentada con el avestruz?', 'nl': 'Welke grote vliegloze vogel is inheems in Zuid-Amerika, verwant aan de struisvogel?', 'sv': 'Vilken stor flygofor fagel ar inhemsk i Sydamerika, slakt med strutsen?', 'fr': "Quel grand oiseau incapable de voler est originaire d'Amerique du Sud, apparente a l'autruche ?", 'da': 'Hvilken stor flyveudygtig fugl er hjemmehørende i Sydamerika, beslægtet med strudsen?'},
 'opt1': {'en': 'Emu', 'de': 'Emu', 'es': 'Emu', 'nl': 'Emoe', 'sv': 'Emu', 'fr': 'Emeu', 'da': 'Emu'},
 'opt2': {'en': 'Rhea', 'de': 'Nandu', 'es': 'Ñandu', 'nl': 'Nandoe', 'sv': 'Nandu', 'fr': 'Nandou', 'da': 'Nandu'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Luang Prabang?', 'de': 'In welchem Land liegt die antike Stadt Luang Prabang?', 'es': '¿En que pais esta la antigua ciudad de Luang Prabang?', 'nl': 'In welk land ligt de oude stad Luang Prabang?', 'sv': 'I vilket land ligger den antika staden Luang Prabang?', 'fr': "Dans quel pays se trouve l'ancienne cite de Luang Prabang ?", 'da': 'Hvilket land huser den antikke by Luang Prabang?'},
 'opt1': {'en': 'Laos', 'de': 'Laos', 'es': 'Laos', 'nl': 'Laos', 'sv': 'Laos', 'fr': 'le Laos', 'da': 'Laos'},
 'opt2': {'en': 'Cambodia', 'de': 'Kambodscha', 'es': 'Camboya', 'nl': 'Cambodja', 'sv': 'Kambodja', 'fr': 'le Cambodge', 'da': 'Cambodja'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of coat, originally worn by fishermen, is short and double-breasted with wide lapels?', 'de': 'Welche Art von Mantel, urspruenglich von Fischern getragen, ist kurz und zweireihig mit breiten Revers?', 'es': '¿Que tipo de abrigo, usado originalmente por pescadores, es corto y cruzado con solapas anchas?', 'nl': 'Welk type jas, oorspronkelijk gedragen door vissers, is kort en dubbelrijs met brede revers?', 'sv': 'Vilken typ av kappa, ursprungligen buren av fiskare, ar kort och dubbelknäppt med breda slag?', 'fr': "Quel type de manteau, porte a l'origine par les pecheurs, est court et croise avec de larges revers ?", 'da': 'Hvilken type frakke, oprindeligt båret af fiskere, er kort og dobbeltradet med brede revers?'},
 'opt1': {'en': 'Peacoat', 'de': 'Pea Coat', 'es': 'Chaqueton marinero', 'nl': 'Pijjekker', 'sv': 'Kavaj-kappa', 'fr': 'Caban', 'da': 'Kabanjakke'},
 'opt2': {'en': 'Trench coat', 'de': 'Trenchcoat', 'es': 'Gabardina', 'nl': 'Trenchcoat', 'sv': 'Trenchcoat', 'fr': 'Trench-coat', 'da': 'Trenchcoat'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which teen drama series is set at a high school in a town built on dark secrets, narrated by a mysterious writer?', 'de': 'Welche Teenie-Drama-Serie spielt an einer High School in einer Stadt voller dunkler Geheimnisse, erzaehlt von einem geheimnisvollen Autor?', 'es': '¿Que serie de drama adolescente se ambienta en un instituto en una ciudad construida sobre oscuros secretos, narrada por un escritor misterioso?', 'nl': 'Welke tienerdramaserie speelt zich af op een middelbare school in een stad vol duistere geheimen, verteld door een mysterieuze schrijver?', 'sv': 'Vilken tonarsdramaserie utspelar sig pa en gymnasieskola i en stad byggd pa morka hemligheter, berattad av en mystisk skribent?', 'fr': "Quelle serie dramatique pour adolescents se deroule dans un lycee d'une ville batie sur de sombres secrets, narree par un ecrivain mysterieux ?", 'da': 'Hvilken teenagedrama-serie foregår på en gymnasieskole i en by bygget på mørke hemmeligheder, fortalt af en mystisk forfatter?'},
 'opt1': {'en': 'Pretty Little Liars', 'de': 'Pretty Little Liars', 'es': 'Pretty Little Liars', 'nl': 'Pretty Little Liars', 'sv': 'Pretty Little Liars', 'fr': 'Pretty Little Liars', 'da': 'Pretty Little Liars'},
 'opt2': {'en': 'Riverdale', 'de': 'Riverdale', 'es': 'Riverdale', 'nl': 'Riverdale', 'sv': 'Riverdale', 'fr': 'Riverdale', 'da': 'Riverdale'}},

{'category': 'Nature', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which flower is known for its star-shaped blooms and strong fragrance, used in perfumes and often climbing on trellises?', 'de': 'Welche Blume ist fuer ihre sternfoermigen Blueten und ihren starken Duft bekannt und wird in Parfuems verwendet, oft an Rankgittern kletternd?', 'es': '¿Que flor es conocida por sus flores en forma de estrella y fuerte fragancia, usada en perfumes y a menudo trepadora en enrejados?', 'nl': 'Welke bloem staat bekend om zijn stervormige bloesems en sterke geur, gebruikt in parfums en klimt vaak op latwerk?', 'sv': 'Vilken blomma ar kand for sina stjarnformade blommor och starka doft, anvands i parfymer och klattrar ofta pa spaljeer?', 'fr': "Quelle fleur est connue pour ses fleurs en etoile et son parfum puissant, utilisee en parfumerie et grimpant souvent sur des treillis ?", 'da': 'Hvilken blomst er kendt for sine stjerneformede blomster og stærke duft, brugt i parfumer og klatrer ofte på gitre?'},
 'opt1': {'en': 'Honeysuckle', 'de': 'Geissblatt', 'es': 'Madreselva', 'nl': 'Kamperfoelie', 'sv': 'Kaprifol', 'fr': 'Chevrefeuille', 'da': 'Kaprifolium'},
 'opt2': {'en': 'Jasmine', 'de': 'Jasmin', 'es': 'Jazmin', 'nl': 'Jasmijn', 'sv': 'Jasmin', 'fr': 'Jasmin', 'da': 'Jasmin'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a deep-fried breaded pork cutlet?', 'de': 'Welches suddstaatliche Gericht ist ein frittiertes, paniertes Schweineschnitzel?', 'es': '¿Que plato sureño de EE.UU. es una chuleta de cerdo empanada y frita?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een gefrituurde gepaneerde varkenskotelet?', 'sv': 'Vilken sydstatlig ratt ar en friterad panerad flaskkotlett?', 'fr': "Quel plat du Sud des États-Unis est une cotelette de porc panee et frite ?", 'da': 'Hvilken sydstatsret er en friturestegt paneret svinekotelet?'},
 'opt1': {'en': 'Fried pork chop', 'de': 'Frittiertes Schweinekotelett', 'es': 'Chuleta de cerdo frita', 'nl': 'Gefrituurde varkenskotelet', 'sv': 'Friterad flaskkotlett', 'fr': 'Cotelette de porc frite', 'da': 'Friturestegt svinekotelet'},
 'opt2': {'en': 'Chicken fried steak', 'de': 'Chicken Fried Steak', 'es': 'Bistec frito estilo pollo', 'nl': 'Chicken fried steak', 'sv': 'Chicken fried steak', 'fr': 'Chicken fried steak', 'da': 'Chicken fried steak'}},

{'category': 'Science', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which vitamin is important for wound healing and found in citrus and peppers?', 'de': 'Welches Vitamin ist wichtig fuer die Wundheilung und in Zitrusfruechten und Paprika enthalten?', 'es': '¿Que vitamina es importante para la cicatrizacion de heridas y se encuentra en citricos y pimientos?', 'nl': 'Welk vitamine is belangrijk voor wondgenezing en te vinden in citrusvruchten en paprika\'s?', 'sv': 'Vilket vitamin ar viktigt for sarlakning och finns i citrus och paprika?', 'fr': "Quelle vitamine est importante pour la cicatrisation des plaies et se trouve dans les agrumes et les poivrons ?", 'da': 'Hvilket vitamin er vigtigt for sårheling og findes i citrus og peberfrugter?'},
 'opt1': {'en': 'Vitamin C', 'de': 'Vitamin C', 'es': 'Vitamina C', 'nl': 'Vitamine C', 'sv': 'Vitamin C', 'fr': 'Vitamine C', 'da': 'Vitamin C'},
 'opt2': {'en': 'Vitamin B12', 'de': 'Vitamin B12', 'es': 'Vitamina B12', 'nl': 'Vitamine B12', 'sv': 'Vitamin B12', 'fr': 'Vitamine B12', 'da': 'Vitamin B12'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which country is home to the ancient city of Vientiane?', 'de': 'In welchem Land liegt die antike Stadt Vientiane?', 'es': '¿En que pais esta la antigua ciudad de Vientian?', 'nl': 'In welk land ligt de oude stad Vientiane?', 'sv': 'I vilket land ligger den antika staden Vientiane?', 'fr': "Dans quel pays se trouve l'ancienne cite de Vientiane ?", 'da': 'Hvilket land huser den antikke by Vientiane?'},
 'opt1': {'en': 'Laos', 'de': 'Laos', 'es': 'Laos', 'nl': 'Laos', 'sv': 'Laos', 'fr': 'le Laos', 'da': 'Laos'},
 'opt2': {'en': 'Myanmar', 'de': 'Myanmar', 'es': 'Myanmar', 'nl': 'Myanmar', 'sv': 'Myanmar', 'fr': 'le Myanmar', 'da': 'Myanmar'}},

{'category': 'History', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which First Lady was known for her work modernizing the White House kitchen and hosting state dinners?', 'de': 'Welche First Lady war bekannt fuer ihre Arbeit an der Modernisierung der Kueche des Weissen Hauses und der Ausrichtung von Staatsdiners?', 'es': '¿Que primera dama fue conocida por su trabajo modernizando la cocina de la Casa Blanca y organizando cenas de estado?', 'nl': 'Welke First Lady stond bekend om haar werk aan het moderniseren van de keuken van het Witte Huis en het organiseren van staatsdiners?', 'sv': 'Vilken First Lady var kand for sitt arbete med att modernisera Vita husets kok och stod for statsmiddagar?', 'fr': "Quelle Première dame etait connue pour avoir modernise la cuisine de la Maison-Blanche et organise des diners d'État ?", 'da': 'Hvilken førstedame var kendt for sit arbejde med at modernisere Det Hvide Hus’ køkken og være vært for statsmiddage?'},
 'opt1': {'en': 'Rosalynn Carter', 'de': 'Rosalynn Carter', 'es': 'Rosalynn Carter', 'nl': 'Rosalynn Carter', 'sv': 'Rosalynn Carter', 'fr': 'Rosalynn Carter', 'da': 'Rosalynn Carter'},
 'opt2': {'en': 'Pat Nixon', 'de': 'Pat Nixon', 'es': 'Pat Nixon', 'nl': 'Pat Nixon', 'sv': 'Pat Nixon', 'fr': 'Pat Nixon', 'da': 'Pat Nixon'}},

{'category': 'Music', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which composer wrote 'The Barber of Seville'?", 'de': "Welcher Komponist schrieb 'Der Barbier von Sevilla'?", 'es': "¿Que compositor escribio 'El barbero de Sevilla'?", 'nl': "Welke componist schreef 'De Barbier van Sevilla'?", 'sv': "Vilken kompositor skrev 'Barberaren i Sevilla'?", 'fr': "Quel compositeur a ecrit 'Le Barbier de Seville' ?", 'da': "Hvilken komponist skrev 'Barberen i Sevilla'?"},
 'opt1': {'en': 'Gaetano Donizetti', 'de': 'Gaetano Donizetti', 'es': 'Gaetano Donizetti', 'nl': 'Gaetano Donizetti', 'sv': 'Gaetano Donizetti', 'fr': 'Gaetano Donizetti', 'da': 'Gaetano Donizetti'},
 'opt2': {'en': 'Gioachino Rossini', 'de': 'Gioachino Rossini', 'es': 'Gioachino Rossini', 'nl': 'Gioachino Rossini', 'sv': 'Gioachino Rossini', 'fr': 'Gioachino Rossini', 'da': 'Gioachino Rossini'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which actress played Cinderella in the 2015 live-action Disney film?', 'de': 'Welche Schauspielerin spielte Cinderella im Disney-Realfilm von 2015?', 'es': '¿Que actriz interpreto a Cenicienta en la pelicula de accion real de Disney de 2015?', 'nl': 'Welke actrice speelde Assepoester in de Disney live-actionfilm uit 2015?', 'sv': 'Vilken skadespelerska spelade Askungen i Disneys spelfilmsversion fran 2015?', 'fr': "Quelle actrice a joue Cendrillon dans le film en prises de vues reelles de Disney en 2015 ?", 'da': 'Hvilken skuespiller spillede Askepot i Disneys 2015 live-action-film?'},
 'opt1': {'en': 'Lily James', 'de': 'Lily James', 'es': 'Lily James', 'nl': 'Lily James', 'sv': 'Lily James', 'fr': 'Lily James', 'da': 'Lily James'},
 'opt2': {'en': 'Emma Watson', 'de': 'Emma Watson', 'es': 'Emma Watson', 'nl': 'Emma Watson', 'sv': 'Emma Watson', 'fr': 'Emma Watson', 'da': 'Emma Watson'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a spiced rum cake soaked in syrup, popular in the Caribbean-influenced South?', 'de': 'Welches suddstaatliche Gericht ist ein gewuerzter Rumkuchen, getraenkt in Sirup, beliebt im karibisch beeinflussten Sueden?', 'es': '¿Que plato sureño de EE.UU. es un pastel de ron especiado empapado en almibar, popular en el sur influenciado por el Caribe?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een gekruide rumcake gedrenkt in siroop, populair in het door het Caribisch gebied beinvloede Zuiden?', 'sv': 'Vilken sydstatlig ratt ar en kryddad rom-kaka indranakt i sirap, popular i den karibiskt paverkade sodern?', 'fr': "Quel plat du Sud des États-Unis est un gateau au rhum epice trempe de sirop, populaire dans le Sud influence par les Caraibes ?", 'da': 'Hvilken sydstatsret er en krydret rom-kage dyppet i sirup, populær i den caribisk-inspirerede syden?'},
 'opt1': {'en': 'Rum cake', 'de': 'Rumkuchen', 'es': 'Pastel de ron', 'nl': 'Rumcake', 'sv': 'Romkaka', 'fr': 'Gateau au rhum', 'da': 'Rombabá'},
 'opt2': {'en': 'Pound cake', 'de': 'Rührkuchen', 'es': 'Bizcocho', 'nl': 'Poundcake', 'sv': 'Sockerkaka', 'fr': 'Quatre-quarts', 'da': 'Sandkage'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which U.S. state is home to the Poconos mountain resort region?', 'de': 'In welchem US-Bundesstaat liegt die Bergresort-Region Poconos?', 'es': '¿En que estado de EE.UU. esta la region turistica de montaña Poconos?', 'nl': 'In welke Amerikaanse staat ligt de berg-resortregio de Poconos?', 'sv': 'I vilken amerikansk delstat ligger bergresortregionen Poconos?', 'fr': "Dans quel État americain se trouve la region de villegiature des Poconos ?", 'da': 'I hvilken amerikansk delstat ligger Poconos-bjergresortregionen?'},
 'opt1': {'en': 'Pennsylvania', 'de': 'Pennsylvania', 'es': 'Pensilvania', 'nl': 'Pennsylvania', 'sv': 'Pennsylvania', 'fr': 'la Pennsylvanie', 'da': 'Pennsylvania'},
 'opt2': {'en': 'New York', 'de': 'New York', 'es': 'Nueva York', 'nl': 'New York', 'sv': 'New York', 'fr': "l'État de New York", 'da': 'New York'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which small mammal is known for its distinctive black mask and ringed tail, often raiding trash cans?', 'de': 'Welches kleine Saeugetier ist fuer seine markante schwarze Maske und den geringelten Schwanz bekannt und pluendert oft Mueltonnen?', 'es': '¿Que pequeño mamifero es conocido por su distintiva mascara negra y cola anillada, a menudo saqueando botes de basura?', 'nl': 'Welk klein zoogdier staat bekend om zijn kenmerkende zwarte masker en geringde staart, en plundert vaak vuilnisbakken?', 'sv': 'Vilket litet daggdjur ar kant for sin markanta svarta mask och ringmarkta svans, plundrar ofta soptunnor?', 'fr': "Quel petit mammifere est connu pour son masque noir distinctif et sa queue annelee, pillant souvent les poubelles ?", 'da': 'Hvilket lille pattedyr er kendt for sin karakteristiske sorte maske og ringmærkede hale, plyndrer ofte skraldespande?'},
 'opt1': {'en': 'Opossum', 'de': 'Opossum', 'es': 'Zarigueya', 'nl': 'Opossum', 'sv': 'Opossum', 'fr': 'Opossum', 'da': 'Opossum'},
 'opt2': {'en': 'Raccoon', 'de': 'Waschbaer', 'es': 'Mapache', 'nl': 'Wasbeer', 'sv': 'Tvattbjorn', 'fr': 'Raton laveur', 'da': 'Vaskebjørn'}},

{'category': 'Culture', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which Thai festival involves floating baskets on water to honor a water goddess?', 'de': 'Welches thailaendische Fest beinhaltet das Schwimmenlassen von Koerben auf dem Wasser zu Ehren einer Wassergoettin?', 'es': '¿Que festival tailandes implica hacer flotar cestas en el agua para honrar a una diosa del agua?', 'nl': 'Welk Thais festival omvat het laten drijven van mandjes op water ter ere van een watergodin?', 'sv': 'Vilken thailandsk festival innebar att flyta korgar pa vattnet for att hedra en vattengudinna?', 'fr': "Quel festival thai implique de faire flotter des paniers sur l'eau pour honorer une deesse de l'eau ?", 'da': 'Hvilken thailandsk festival involverer at lade kurve flyde på vand for at ære en vandgudinde?'},
 'opt1': {'en': 'Songkran', 'de': 'Songkran', 'es': 'Songkran', 'nl': 'Songkran', 'sv': 'Songkran', 'fr': 'Songkran', 'da': 'Songkran'},
 'opt2': {'en': 'Loy Krathong', 'de': 'Loy Krathong', 'es': 'Loy Krathong', 'nl': 'Loy Krathong', 'sv': 'Loy Krathong', 'fr': 'Loy Krathong', 'da': 'Loy Krathong'}},

{'category': 'Literature', 'tier': 2, 'correct_index': 2,
 'question': {'en': "Which author wrote 'The Grapes of Wrath'?", 'de': "Welcher Autor schrieb 'Fruechte des Zorns'?", 'es': "¿Que autor escribio 'Las uvas de la ira'?", 'nl': "Welke auteur schreef 'De druiven der gramschap'?", 'sv': "Vilken forfattare skrev 'Vredens druvor'?", 'fr': "Quel auteur a ecrit 'Les Raisins de la colere' ?", 'da': "Hvilken forfatter skrev 'Vredens druer'?"},
 'opt1': {'en': 'Ernest Hemingway', 'de': 'Ernest Hemingway', 'es': 'Ernest Hemingway', 'nl': 'Ernest Hemingway', 'sv': 'Ernest Hemingway', 'fr': 'Ernest Hemingway', 'da': 'Ernest Hemingway'},
 'opt2': {'en': 'John Steinbeck', 'de': 'John Steinbeck', 'es': 'John Steinbeck', 'nl': 'John Steinbeck', 'sv': 'John Steinbeck', 'fr': 'John Steinbeck', 'da': 'John Steinbeck'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Battambang?', 'de': 'In welchem Land liegt die antike Stadt Battambang?', 'es': '¿En que pais esta la antigua ciudad de Battambang?', 'nl': 'In welk land ligt de oude stad Battambang?', 'sv': 'I vilket land ligger den antika staden Battambang?', 'fr': "Dans quel pays se trouve l'ancienne cite de Battambang ?", 'da': 'Hvilket land huser den antikke by Battambang?'},
 'opt1': {'en': 'Laos', 'de': 'Laos', 'es': 'Laos', 'nl': 'Laos', 'sv': 'Laos', 'fr': 'le Laos', 'da': 'Laos'},
 'opt2': {'en': 'Cambodia', 'de': 'Kambodscha', 'es': 'Camboya', 'nl': 'Cambodja', 'sv': 'Kambodja', 'fr': 'le Cambodge', 'da': 'Cambodja'}},

{'category': 'Music', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which instrument is central to Indian classical music, a plucked string instrument with a long neck?', 'de': 'Welches Instrument ist zentral fuer indische klassische Musik, ein gezupftes Saiteninstrument mit langem Hals?', 'es': '¿Que instrumento es central en la musica clasica india, un instrumento de cuerda punteada de mastil largo?', 'nl': 'Welk instrument is centraal in de Indiase klassieke muziek, een getokkeld snaarinstrument met een lange hals?', 'sv': 'Vilket instrument ar centralt i indisk klassisk musik, ett knappt strakinstrument med lang hals?', 'fr': "Quel instrument est central dans la musique classique indienne, un instrument a cordes pincees au long manche ?", 'da': 'Hvilket instrument er centralt i indisk klassisk musik, et klimpret strengeinstrument med lang hals?'},
 'opt1': {'en': 'Sitar', 'de': 'Sitar', 'es': 'Sitar', 'nl': 'Sitar', 'sv': 'Sitar', 'fr': 'Sitar', 'da': 'Sitar'},
 'opt2': {'en': 'Sarod', 'de': 'Sarod', 'es': 'Sarod', 'nl': 'Sarod', 'sv': 'Sarod', 'fr': 'Sarod', 'da': 'Sarod'}},

{'category': 'Food', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which Southern U.S. dish is a deep-fried breaded fish fillet, often catfish?', 'de': 'Welches suddstaatliche Gericht ist ein frittiertes, paniertes Fischfilet, oft Wels?', 'es': '¿Que plato sureño de EE.UU. es un filete de pescado empanado y frito, a menudo bagre?', 'nl': 'Welk zuidelijk Amerikaans gerecht is een gefrituurd gepaneerd visfilet, vaak meerval?', 'sv': 'Vilken sydstatlig ratt ar en friterad panerad fiskfile, ofta malformig fisk (catfish)?', 'fr': "Quel plat du Sud des États-Unis est un filet de poisson pane et frit, souvent du poisson-chat ?", 'da': 'Hvilken sydstatsret er en friturestegt paneret fiskefilet, ofte malle?'},
 'opt1': {'en': 'Fried catfish', 'de': 'Frittierter Wels', 'es': 'Bagre frito', 'nl': 'Gefrituurde meerval', 'sv': 'Friterad malformig fisk', 'fr': 'Poisson-chat frit', 'da': 'Friturestegt malle'},
 'opt2': {'en': 'Fish and chips', 'de': 'Fish and Chips', 'es': 'Pescado con papas fritas', 'nl': 'Fish and chips', 'sv': 'Fish and chips', 'fr': 'Fish and chips', 'da': 'Fish and chips'}},

{'category': 'Science', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which mineral, found in whole grains, supports muscle and nerve function?', 'de': 'Welches Mineral, in Vollkorngetreide enthalten, unterstuetzt Muskel- und Nervenfunktion?', 'es': '¿Que mineral, presente en granos integrales, apoya la funcion muscular y nerviosa?', 'nl': 'Welk mineraal, aanwezig in volkoren granen, ondersteunt spier- en zenuwfunctie?', 'sv': 'Vilket mineral, som finns i fullkorn, stodjer muskel- och nervfunktion?', 'fr': "Quel mineral, present dans les cereales completes, soutient la fonction musculaire et nerveuse ?", 'da': 'Hvilket mineral, fundet i fuldkorn, understøtter muskel- og nervefunktion?'},
 'opt1': {'en': 'Potassium', 'de': 'Kalium', 'es': 'Potasio', 'nl': 'Kalium', 'sv': 'Kalium', 'fr': 'Potassium', 'da': 'Kalium'},
 'opt2': {'en': 'Magnesium', 'de': 'Magnesium', 'es': 'Magnesio', 'nl': 'Magnesium', 'sv': 'Magnesium', 'fr': 'Magnesium', 'da': 'Magnesium'}},

{'category': 'Entertainment', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which film has Sandra Bullock playing an astronaut stranded in space?', 'de': 'In welchem Film spielt Sandra Bullock eine Astronautin, die im All gestrandet ist?', 'es': '¿En que pelicula Sandra Bullock interpreta a una astronauta varada en el espacio?', 'nl': 'In welke film speelt Sandra Bullock een astronaut die gestrand is in de ruimte?', 'sv': 'I vilken film spelar Sandra Bullock en astronaut som strandat i rymden?', 'fr': "Dans quel film Sandra Bullock joue-t-elle une astronaute coincee dans l'espace ?", 'da': 'I hvilken film spiller Sandra Bullock en astronaut strandet i rummet?'},
 'opt1': {'en': 'Interstellar', 'de': 'Interstellar', 'es': 'Interstellar', 'nl': 'Interstellar', 'sv': 'Interstellar', 'fr': 'Interstellar', 'da': 'Interstellar'},
 'opt2': {'en': 'Gravity', 'de': 'Gravity', 'es': 'Gravity', 'nl': 'Gravity', 'sv': 'Gravity', 'fr': 'Gravity', 'da': 'Gravity'}},

{'category': 'Geography', 'tier': 2, 'correct_index': 2,
 'question': {'en': 'Which country is home to the ancient city of Siem Reap, gateway to Angkor Wat?', 'de': 'In welchem Land liegt die antike Stadt Siem Reap, Tor zu Angkor Wat?', 'es': '¿En que pais esta la antigua ciudad de Siem Reap, puerta de entrada a Angkor Wat?', 'nl': 'In welk land ligt de oude stad Siem Reap, poort naar Angkor Wat?', 'sv': 'I vilket land ligger den antika staden Siem Reap, port till Angkor Wat?', 'fr': "Dans quel pays se trouve l'ancienne cite de Siem Reap, porte d'entree d'Angkor Wat ?", 'da': 'Hvilket land huser den antikke by Siem Reap, port til Angkor Wat?'},
 'opt1': {'en': 'Vietnam', 'de': 'Vietnam', 'es': 'Vietnam', 'nl': 'Vietnam', 'sv': 'Vietnam', 'fr': 'le Vietnam', 'da': 'Vietnam'},
 'opt2': {'en': 'Cambodia', 'de': 'Kambodscha', 'es': 'Camboya', 'nl': 'Cambodja', 'sv': 'Kambodja', 'fr': 'le Cambodge', 'da': 'Cambodja'}},

{'category': 'Fashion', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which type of jewelry is worn on the ear and can be a stud, hoop, or dangle style?', 'de': 'Welche Art von Schmuck wird am Ohr getragen und kann als Stecker, Creole oder haengender Ohrring sein?', 'es': '¿Que tipo de joya se lleva en la oreja y puede ser de estilo boton, aro o colgante?', 'nl': 'Welk type sieraad wordt aan het oor gedragen en kan een stud, hoepel, of hangende stijl zijn?', 'sv': 'Vilken typ av smycke baers pa orat och kan vara knapp-, ring- eller hangande stil?', 'fr': "Quel type de bijou se porte a l'oreille et peut etre en clou, en anneau ou en pendant ?", 'da': 'Hvilken type smykke bæres på øret og kan være i knap-, ring- eller hængende stil?'},
 'opt1': {'en': 'Earrings', 'de': 'Ohrringe', 'es': 'Aretes', 'nl': 'Oorbellen', 'sv': 'Örhängen', 'fr': "Boucles d'oreilles", 'da': 'Øreringe'},
 'opt2': {'en': 'Cufflinks', 'de': 'Manschettenknoepfe', 'es': 'Mancuernillas', 'nl': 'Manchetknopen', 'sv': 'Manschettknappar', 'fr': 'Boutons de manchette', 'da': 'Manchetknapper'}},

{'category': 'Animals', 'tier': 2, 'correct_index': 1,
 'question': {'en': 'Which large bird has a distinctive throat pouch and dives for fish from great heights?', 'de': 'Welcher grosse Vogel hat einen markanten Kehlsack und stuerzt sich aus grosser Hoehe auf Fische?', 'es': '¿Que gran ave tiene una distintiva bolsa en la garganta y se zambulle por peces desde grandes alturas?', 'nl': 'Welke grote vogel heeft een kenmerkende keelzak en duikt van grote hoogte naar vis?', 'sv': 'Vilken stor fagel har en markant strupepase och dyker efter fisk fran stora hojder?', 'fr': "Quel grand oiseau a une poche gulaire distinctive et plonge pour attraper des poissons depuis de grandes hauteurs ?", 'da': 'Hvilken stor fugl har en karakteristisk strubepose og dykker efter fisk fra store højder?'},
 'opt1': {'en': 'Pelican', 'de': 'Pelikan', 'es': 'Pelicano', 'nl': 'Pelikaan', 'sv': 'Pelikan', 'fr': 'Pelican', 'da': 'Pelikan'},
 'opt2': {'en': 'Gannet', 'de': 'Basstoelpel', 'es': 'Alcatraz', 'nl': 'Jan-van-gent', 'sv': 'Havssula', 'fr': 'Fou de Bassan', 'da': 'Sule'}},

{'category': 'History', 'tier': 2, 'correct_index': 1,
 'question': {'en': "Which First Lady wrote the memoir 'Living History' about her time in the White House?", 'de': "Welche First Lady schrieb die Memoiren 'Living History' ueber ihre Zeit im Weissen Haus?", 'es': "¿Que primera dama escribio las memorias 'Living History' sobre su tiempo en la Casa Blanca?", 'nl': "Welke First Lady schreef de memoires 'Living History' over haar tijd in het Witte Huis?", 'sv': "Vilken First Lady skrev memoarerna 'Living History' om sin tid i Vita huset?", 'fr': "Quelle Première dame a ecrit les memoires 'Living History' sur son temps a la Maison-Blanche ?", 'da': 'Hvilken førstedame skrev erindringerne "Living History" om sin tid i Det Hvide Hus?'},
 'opt1': {'en': 'Hillary Clinton', 'de': 'Hillary Clinton', 'es': 'Hillary Clinton', 'nl': 'Hillary Clinton', 'sv': 'Hillary Clinton', 'fr': 'Hillary Clinton', 'da': 'Hillary Clinton'},
 'opt2': {'en': 'Laura Bush', 'de': 'Laura Bush', 'es': 'Laura Bush', 'nl': 'Laura Bush', 'sv': 'Laura Bush', 'fr': 'Laura Bush', 'da': 'Laura Bush'}},

{'category': 'History', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which First Lady was a professional singer who trained in Europe before her marriage?', 'de': 'Welche First Lady war eine professionelle Saengerin, die vor ihrer Ehe in Europa ausgebildet wurde?', 'es': '¿Que primera dama fue una cantante profesional que se formo en Europa antes de su matrimonio?', 'nl': 'Welke First Lady was een professionele zangeres die voor haar huwelijk in Europa was opgeleid?', 'sv': 'Vilken First Lady var en professionell sangerska som utbildade sig i Europa innan sitt aktenskap?', 'fr': "Quelle Première dame etait une chanteuse professionnelle formee en Europe avant son mariage ?", 'da': 'Hvilken førstedame var en professionel sanger, der trænede i Europa før sit ægteskab?'},
 'opt1': {'en': 'Helen Taft', 'de': 'Helen Taft', 'es': 'Helen Taft', 'nl': 'Helen Taft', 'sv': 'Helen Taft', 'fr': 'Helen Taft', 'da': 'Helen Taft'},
 'opt2': {'en': 'Grace Coolidge', 'de': 'Grace Coolidge', 'es': 'Grace Coolidge', 'nl': 'Grace Coolidge', 'sv': 'Grace Coolidge', 'fr': 'Grace Coolidge', 'da': 'Grace Coolidge'}},

{'category': 'Literature', 'tier': 3, 'correct_index': 2,
 'question': {'en': "Which South African author wrote 'Disgrace', winning the Booker Prize?", 'de': "Welcher suedafrikanische Autor schrieb 'Schande' und gewann den Booker Prize?", 'es': "¿Que autor sudafricano escribio 'Desgracia', ganando el Premio Booker?", 'nl': "Welke Zuid-Afrikaanse auteur schreef 'Onteerd', winnaar van de Booker Prize?", 'sv': "Vilken sydafrikansk forfattare skrev 'Vanara', som vann Bookerpriset?", 'fr': "Quel auteur sud-africain a ecrit 'Disgrace', remportant le prix Booker ?", 'da': "Hvilken sydafrikansk forfatter skrev 'Vanære', som vandt Booker-prisen?"},
 'opt1': {'en': 'Nadine Gordimer', 'de': 'Nadine Gordimer', 'es': 'Nadine Gordimer', 'nl': 'Nadine Gordimer', 'sv': 'Nadine Gordimer', 'fr': 'Nadine Gordimer', 'da': 'Nadine Gordimer'},
 'opt2': {'en': 'J.M. Coetzee', 'de': 'J.M. Coetzee', 'es': 'J.M. Coetzee', 'nl': 'J.M. Coetzee', 'sv': 'J.M. Coetzee', 'fr': 'J.M. Coetzee', 'da': 'J.M. Coetzee'}},

{'category': 'Science', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which vitamin, found in nuts and seeds, acts as an antioxidant protecting cells?', 'de': 'Welches Vitamin, in Nuessen und Samen enthalten, wirkt als Antioxidans zum Schutz der Zellen?', 'es': '¿Que vitamina, presente en frutos secos y semillas, actua como antioxidante protegiendo las celulas?', 'nl': 'Welk vitamine, aanwezig in noten en zaden, werkt als antioxidant om cellen te beschermen?', 'sv': 'Vilket vitamin, som finns i notter och fron, fungerar som en antioxidant som skyddar celler?', 'fr': "Quelle vitamine, presente dans les noix et les graines, agit comme antioxydant protegeant les cellules ?", 'da': 'Hvilket vitamin, fundet i nødder og frø, fungerer som en antioxidant, der beskytter celler?'},
 'opt1': {'en': 'Vitamin E', 'de': 'Vitamin E', 'es': 'Vitamina E', 'nl': 'Vitamine E', 'sv': 'Vitamin E', 'fr': 'Vitamine E', 'da': 'Vitamin E'},
 'opt2': {'en': 'Vitamin K', 'de': 'Vitamin K', 'es': 'Vitamina K', 'nl': 'Vitamine K', 'sv': 'Vitamin K', 'fr': 'Vitamine K', 'da': 'Vitamin K'}},

{'category': 'Arts', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which German Renaissance artist created detailed woodcuts and engravings like 'Melencolia I'?", 'de': "Welcher deutsche Renaissance-Kuenstler schuf detaillierte Holzschnitte und Stiche wie 'Melencolia I'?", 'es': "¿Que artista aleman del Renacimiento creo grabados y xilografias detallados como 'Melencolia I'?", 'nl': "Welke Duitse renaissancekunstenaar creëerde gedetailleerde houtsneden en gravures zoals 'Melencolia I'?", 'sv': "Vilken tysk renässanskonstnar skapade detaljerade traesnitt och gravyrer som 'Melencolia I'?", 'fr': "Quel artiste allemand de la Renaissance a cree des gravures sur bois et sur cuivre detaillees comme 'Melancolia I' ?", 'da': "Hvilken tysk renæssancekunstner skabte detaljerede træsnit og gravurer som 'Melencolia I'?"},
 'opt1': {'en': 'Albrecht Durer', 'de': 'Albrecht Duerer', 'es': 'Alberto Durero', 'nl': 'Albrecht Durer', 'sv': 'Albrecht Durer', 'fr': 'Albrecht Durer', 'da': 'Albrecht Dürer'},
 'opt2': {'en': 'Hans Holbein', 'de': 'Hans Holbein', 'es': 'Hans Holbein', 'nl': 'Hans Holbein', 'sv': 'Hans Holbein', 'fr': 'Hans Holbein', 'da': 'Hans Holbein'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which Pacific island nation is made up of over 300 islands and was formerly known as the Friendly Islands?', 'de': 'Welcher pazifische Inselstaat besteht aus ueber 300 Inseln und war frueher als Freundschaftsinseln bekannt?', 'es': '¿Que nacion insular del Pacifico se compone de mas de 300 islas y era conocida anteriormente como las Islas Amistad?', 'nl': 'Welke Pacifische eilandstaat bestaat uit meer dan 300 eilanden en stond voorheen bekend als de Vriendschapseilanden?', 'sv': 'Vilken oceansk oestat bestar av over 300 oar och var tidigare kand som Vanskapsoarna?', 'fr': "Quelle nation insulaire du Pacifique est composee de plus de 300 iles et etait autrefois connue sous le nom d'iles des Amis ?", 'da': 'Hvilken stillehavsøstat består af over 300 øer og var tidligere kendt som De Venlige Øer?'},
 'opt1': {'en': 'Fiji', 'de': 'Fidschi', 'es': 'Fiyi', 'nl': 'Fiji', 'sv': 'Fiji', 'fr': 'les Fidji', 'da': 'Fiji'},
 'opt2': {'en': 'Tonga', 'de': 'Tonga', 'es': 'Tonga', 'nl': 'Tonga', 'sv': 'Tonga', 'fr': 'les Tonga', 'da': 'Tonga'}},

{'category': 'Music', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Hungarian composer wrote 'Hungarian Rhapsodies' for piano?", 'de': "Welcher ungarische Komponist schrieb die 'Ungarischen Rhapsodien' fuer Klavier?", 'es': "¿Que compositor hungaro escribio las 'Rapsodias hungaras' para piano?", 'nl': "Welke Hongaarse componist schreef de 'Hongaarse Rapsodieen' voor piano?", 'sv': "Vilken ungersk kompositor skrev 'Ungerska rapsodier' for piano?", 'fr': "Quel compositeur hongrois a ecrit les 'Rhapsodies hongroises' pour piano ?", 'da': "Hvilken ungarsk komponist skrev 'Ungarske rhapsodier' for klaver?"},
 'opt1': {'en': 'Franz Liszt', 'de': 'Franz Liszt', 'es': 'Franz Liszt', 'nl': 'Franz Liszt', 'sv': 'Franz Liszt', 'fr': 'Franz Liszt', 'da': 'Franz Liszt'},
 'opt2': {'en': 'Bela Bartok', 'de': 'Bela Bartok', 'es': 'Bela Bartok', 'nl': 'Bela Bartok', 'sv': 'Bela Bartok', 'fr': 'Bela Bartok', 'da': 'Bela Bartok'}},

{'category': 'Food', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which Spanish dish consists of small savory dishes served as appetizers?', 'de': 'Welches spanische Gericht besteht aus kleinen herzhaften Gerichten, die als Vorspeisen serviert werden?', 'es': '¿Que plato español consiste en pequeños platos salados servidos como aperitivos?', 'nl': 'Welk Spaans gerecht bestaat uit kleine hartige gerechten geserveerd als voorgerecht?', 'sv': 'Vilken spansk ratt bestar av sma salta ratter serverade som forratter?', 'fr': "Quel plat espagnol consiste en de petits plats sales servis en apero ?", 'da': 'Hvilken spansk ret består af små salte retter serveret som forretter?'},
 'opt1': {'en': 'Tapas', 'de': 'Tapas', 'es': 'Tapas', 'nl': 'Tapas', 'sv': 'Tapas', 'fr': 'Tapas', 'da': 'Tapas'},
 'opt2': {'en': 'Pintxos', 'de': 'Pintxos', 'es': 'Pintxos', 'nl': 'Pintxos', 'sv': 'Pintxos', 'fr': 'Pintxos', 'da': 'Pintxos'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'In Egyptian mythology, which god rules the underworld and judges the dead?', 'de': 'In der aegyptischen Mythologie, welcher Gott herrscht ueber die Unterwelt und richtet die Toten?', 'es': '¿En la mitologia egipcia, que dios gobierna el inframundo y juzga a los muertos?', 'nl': 'In de Egyptische mythologie, welke god heerst over de onderwereld en oordeelt over de doden?', 'sv': 'I egyptisk mytologi, vilken gud harskar over underjorden och domer de doda?', 'fr': "Dans la mythologie egyptienne, quel dieu regne sur les enfers et juge les morts ?", 'da': 'I egyptisk mytologi, hvilken gud hersker over underverdenen og dømmer de døde?'},
 'opt1': {'en': 'Anubis', 'de': 'Anubis', 'es': 'Anubis', 'nl': 'Anubis', 'sv': 'Anubis', 'fr': 'Anubis', 'da': 'Anubis'},
 'opt2': {'en': 'Osiris', 'de': 'Osiris', 'es': 'Osiris', 'nl': 'Osiris', 'sv': 'Osiris', 'fr': 'Osiris', 'da': 'Osiris'}},

{'category': 'Entertainment', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which actress won an Oscar for playing a British monarch in 'The Favourite'?", 'de': "Welche Schauspielerin gewann einen Oscar fuer ihre Rolle als britische Monarchin in 'The Favourite'?", 'es': "¿Que actriz gano un Oscar por interpretar a una monarca britanica en 'La favorita'?", 'nl': "Welke actrice won een Oscar voor haar rol als Britse vorstin in 'The Favourite'?", 'sv': "Vilken skadespelerska vann en Oscar for att ha spelat en brittisk monark i 'The Favourite'?", 'fr': "Quelle actrice a remporte un Oscar pour son role de monarque britannique dans 'La Favorite' ?", 'da': "Hvilken skuespiller vandt en Oscar for at spille en britisk monark i 'The Favourite'?"},
 'opt1': {'en': 'Olivia Colman', 'de': 'Olivia Colman', 'es': 'Olivia Colman', 'nl': 'Olivia Colman', 'sv': 'Olivia Colman', 'fr': 'Olivia Colman', 'da': 'Olivia Colman'},
 'opt2': {'en': 'Emma Stone', 'de': 'Emma Stone', 'es': 'Emma Stone', 'nl': 'Emma Stone', 'sv': 'Emma Stone', 'fr': 'Emma Stone', 'da': 'Emma Stone'}},

{'category': 'Space', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which asteroid belt lies between the orbits of Mars and Jupiter?', 'de': 'Welcher Asteroidenguertel liegt zwischen den Umlaufbahnen von Mars und Jupiter?', 'es': '¿Que cinturon de asteroides se encuentra entre las orbitas de Marte y Jupiter?', 'nl': 'Welke asteroidengordel ligt tussen de banen van Mars en Jupiter?', 'sv': 'Vilket asteroidbälte ligger mellan Mars och Jupiters banor?', 'fr': "Quelle ceinture d'asteroides se trouve entre les orbites de Mars et Jupiter ?", 'da': 'Hvilket asteroidebælte ligger mellem Mars’ og Jupiters baner?'},
 'opt1': {'en': 'Kuiper belt', 'de': 'Kuipterguertel', 'es': 'Cinturon de Kuiper', 'nl': 'Kuipergordel', 'sv': 'Kuiperbältet', 'fr': 'Ceinture de Kuiper', 'da': 'Kuiperbæltet'},
 'opt2': {'en': 'Main asteroid belt', 'de': 'Hauptasteroidenguertel', 'es': 'Cinturon principal de asteroides', 'nl': 'Hoofdasteroidengordel', 'sv': 'Huvudasteroidbältet', 'fr': "Ceinture principale d'asteroides", 'da': 'Hovedasteroidebælte'}},

{'category': 'Geography', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which West African country is home to the ancient city of Timbuktu?', 'de': 'In welchem westafrikanischen Land liegt die antike Stadt Timbuktu?', 'es': '¿En que pais de Africa occidental esta la antigua ciudad de Tombuctu?', 'nl': 'In welk West-Afrikaans land ligt de oude stad Timboektoe?', 'sv': 'I vilket vastafrikanskt land ligger den antika staden Timbuktu?', 'fr': "Dans quel pays d'Afrique de l'Ouest se trouve l'ancienne cite de Tombouctou ?", 'da': 'Hvilket vestafrikansk land huser den antikke by Timbuktu?'},
 'opt1': {'en': 'Senegal', 'de': 'Senegal', 'es': 'Senegal', 'nl': 'Senegal', 'sv': 'Senegal', 'fr': 'le Senegal', 'da': 'Senegal'},
 'opt2': {'en': 'Mali', 'de': 'Mali', 'es': 'Mali', 'nl': 'Mali', 'sv': 'Mali', 'fr': 'le Mali', 'da': 'Mali'}},

{'category': 'Fashion', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'Which British fashion house is known for its trench coats and check pattern lining?', 'de': 'Welches britische Modehaus ist fuer seine Trenchcoats und Karomuster-Futter bekannt?', 'es': '¿Que casa de moda britanica es conocida por sus gabardinas y forro de patron de cuadros?', 'nl': 'Welk Brits modehuis staat bekend om zijn trenchcoats en ruitpatroonvoering?', 'sv': 'Vilket brittiskt modehus ar kant for sina trenchcoats och rutmonstrade foder?', 'fr': "Quelle maison de mode britannique est connue pour ses trench-coats et sa doublure a carreaux ?", 'da': 'Hvilket britisk modehus er kendt for sine trenchcoats og ternet foring?'},
 'opt1': {'en': 'Burberry', 'de': 'Burberry', 'es': 'Burberry', 'nl': 'Burberry', 'sv': 'Burberry', 'fr': 'Burberry', 'da': 'Burberry'},
 'opt2': {'en': 'Aquascutum', 'de': 'Aquascutum', 'es': 'Aquascutum', 'nl': 'Aquascutum', 'sv': 'Aquascutum', 'fr': 'Aquascutum', 'da': 'Aquascutum'}},

{'category': 'Philosophy', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which Greek philosopher wrote 'The Republic', describing an ideal state?", 'de': "Welcher griechische Philosoph schrieb 'Der Staat', der einen idealen Staat beschreibt?", 'es': "¿Que filosofo griego escribio 'La Republica', describiendo un estado ideal?", 'nl': "Welke Griekse filosoof schreef 'De Staat', waarin een ideale staat wordt beschreven?", 'sv': "Vilken grekisk filosof skrev 'Staten', som beskriver en ideal stat?", 'fr': "Quel philosophe grec a ecrit 'La Republique', decrivant un État ideal ?", 'da': "Hvilken græsk filosof skrev 'Staten', der beskriver en ideel stat?"},
 'opt1': {'en': 'Plato', 'de': 'Platon', 'es': 'Platon', 'nl': 'Plato', 'sv': 'Platon', 'fr': 'Platon', 'da': 'Platon'},
 'opt2': {'en': 'Aristotle', 'de': 'Aristoteles', 'es': 'Aristoteles', 'nl': 'Aristoteles', 'sv': 'Aristoteles', 'fr': 'Aristote', 'da': 'Aristoteles'}},

{'category': 'History', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which quilting pattern uses interlocking circles resembling a wedding ring, often in pastel fabrics?', 'de': 'Welches Quiltmuster verwendet ineinandergreifende Kreise, die einem Ehering aehneln, oft in Pastellstoffen?', 'es': '¿Que patron de acolchado usa circulos entrelazados que se asemejan a una alianza de boda, a menudo en telas pastel?', 'nl': 'Welk quiltpatroon gebruikt in elkaar grijpende cirkels die op een trouwring lijken, vaak in pasteltinten?', 'sv': 'Vilket lapptacksmonster anvander sammanlankade cirklar som liknar en vigselring, ofta i pastellfargat tyg?', 'fr': "Quel motif de patchwork utilise des cercles entrelaces ressemblant a une alliance, souvent dans des tissus pastel ?", 'da': 'Hvilket quiltmønster bruger sammenkoblede cirkler, der ligner en vielsesring, ofte i pastelfarvede stoffer?'},
 'opt1': {'en': 'Trip Around the World', 'de': 'Reise um die Welt', 'es': 'Vuelta al mundo', 'nl': 'Reis om de Wereld', 'sv': 'Resa jorden runt', 'fr': 'Tour du monde', 'da': 'Rejse rundt om verden'},
 'opt2': {'en': 'Double Wedding Ring', 'de': 'Doppelter Ehering', 'es': 'Doble anillo de boda', 'nl': 'Dubbele Trouwring', 'sv': 'Dubbel vigselring', 'fr': 'Double anneau de mariage', 'da': 'Dobbelt vielsesring'}},

{'category': 'Nature', 'tier': 3, 'correct_index': 1,
 'question': {'en': "Which flowering plant's roots are used to make traditional root beer?", 'de': 'Welche bluehende Pflanze liefert die Wurzeln fuer traditionelles Root Beer?', 'es': '¿Que planta con flores tiene raices usadas para hacer root beer tradicional?', 'nl': 'Welke bloeiende plant levert de wortels voor traditionele root beer?', 'sv': 'Vilken blommande vaxts rotter anvands for att gora traditionell root beer?', 'fr': "Quelle plante a fleurs a des racines utilisees pour faire la root beer traditionnelle ?", 'da': 'Hvilken blomstrende plantes rødder bruges til at lave traditionel root beer?'},
 'opt1': {'en': 'Sassafras', 'de': 'Sassafras', 'es': 'Sasafras', 'nl': 'Sassafras', 'sv': 'Sassafras', 'fr': 'Sassafras', 'da': 'Sassafras'},
 'opt2': {'en': 'Sarsaparilla', 'de': 'Sarsaparille', 'es': 'Zarzaparrilla', 'nl': 'Sarsaparilla', 'sv': 'Sarsaparilla', 'fr': 'Salsepareille', 'da': 'Sarsaparilla'}},

{'category': 'Architecture', 'tier': 3, 'correct_index': 2,
 'question': {'en': 'Which architect designed the Pompidou Centre in Paris with its exposed pipes and structure?', 'de': 'Welcher Architekt entwarf das Centre Pompidou in Paris mit seinen freiliegenden Rohren und Struktur?', 'es': '¿Que arquitecto diseño el Centro Pompidou en Paris con sus tuberias y estructura expuestas?', 'nl': 'Welke architect ontwierp het Centre Pompidou in Parijs met zijn blootliggende leidingen en structuur?', 'sv': 'Vilken arkitekt designade Centre Pompidou i Paris med sina synliga ror och struktur?', 'fr': "Quel architecte a conçu le Centre Pompidou a Paris avec ses tuyaux et sa structure apparents ?", 'da': 'Hvilken arkitekt designede Centre Pompidou i Paris med dets synlige rør og struktur?'},
 'opt1': {'en': 'Jean Nouvel', 'de': 'Jean Nouvel', 'es': 'Jean Nouvel', 'nl': 'Jean Nouvel', 'sv': 'Jean Nouvel', 'fr': 'Jean Nouvel', 'da': 'Jean Nouvel'},
 'opt2': {'en': 'Renzo Piano and Richard Rogers', 'de': 'Renzo Piano und Richard Rogers', 'es': 'Renzo Piano y Richard Rogers', 'nl': 'Renzo Piano en Richard Rogers', 'sv': 'Renzo Piano och Richard Rogers', 'fr': 'Renzo Piano et Richard Rogers', 'da': 'Renzo Piano og Richard Rogers'}},

{'category': 'Mythology', 'tier': 3, 'correct_index': 1,
 'question': {'en': 'In Egyptian mythology, which falcon-headed god is associated with the sky and kingship?', 'de': 'In der aegyptischen Mythologie, welcher Falkengott wird mit dem Himmel und dem Koenigtum assoziiert?', 'es': '¿En la mitologia egipcia, que dios con cabeza de halcon se asocia con el cielo y la realeza?', 'nl': 'In de Egyptische mythologie, welke god met valkenhoofd wordt geassocieerd met de hemel en koningschap?', 'sv': 'I egyptisk mytologi, vilken falkhuvad gud forknippas med himlen och kungadomet?', 'fr': "Dans la mythologie egyptienne, quel dieu a tete de faucon est associe au ciel et a la royaute ?", 'da': 'I egyptisk mytologi, hvilken falkehovedet gud forbindes med himlen og kongedømmet?'},
 'opt1': {'en': 'Horus', 'de': 'Horus', 'es': 'Horus', 'nl': 'Horus', 'sv': 'Horus', 'fr': 'Horus', 'da': 'Horus'},
 'opt2': {'en': 'Ra', 'de': 'Ra', 'es': 'Ra', 'nl': 'Ra', 'sv': 'Ra', 'fr': 'Ra', 'da': 'Ra'}},
]
