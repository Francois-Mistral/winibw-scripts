'Script permettant d'interagir avec IdRef :

'-Récupération du type de zone et des valeurs $a et $b de la ligne où se trouve le curseur en édition de notice WinIBW
'-Sélection automatique de l'index IdRef correspondant au type de zone et présaisie des valeurs $a et $b dans la zone de recherche IdRef
'-Récupération des informations bibliographiques pour faciliter une éventuelle création de notice d'autorité dans IdRef
'-Authentification automatique avec le login WinIBW et le mot de passe saisie dans ce script (variable motdepasse à renseigner)
'-Lancement automatique de la recherche dans IdRef 
	
Dim IEDoc
Dim IE  		

Sub IdRef()

	' Contient le login WinIBW/IdRef
	utilisateur = application.activeWindow.Variable("P3GUK")
	' Mot de passe a renseigner si on souhaite une authentification automatique dans IdRef
	motdepasse = ""

	'End of (Declarations)
	set IE = nothing
    set shapp=createobject("shell.application")
     Dim InputTexte
	'MsgBox  "==>" + IE.Visible
    on error resume next
    'pour ouvrir si pas ouvert
    For Each owin In shapp.Windows
         if left(owin.document.location.href,len("https://www.idref.fr"))="https://www.idref.fr" then
            if err.number = 0 then
                    set IE = owin
                    'MsgBox "ok"
              end if
        end if
    err.clear
    Next

    on error goto 0
    if IE is nothing then
        'MsgBox  "Window Not Open"
         Set IE = CreateObject("InternetExplorer.Application")
    end if

    IE.Navigate2 "https://www.idref.fr"    
	Do While IE.readystate <> 4  
    Loop  
    Set IEDoc = IE.document
      
	'Selectionne la ligne ou se trouve le curseur
    application.ActiveWindow.title.EndOfField()
	application.ActiveWindow.title.StartOfField(True)    
	  
	'sauvegarde la ligne où se trouve le curseur. Afin de pouvoir remettre le focus avant le rapatriement du ppn de l'autorité
	ligneSelection = ""
	ligneSelection = application.ActiveWindow.title.selection()
	
	' si la selection est trop courte, on affiche un message d'erreur
	if len(ligneSelection) < 10  then
		MsgBox  "Mauvaise sélection : pas assez d'informations"
	else

	Set LogOut = IEDoc.getElementById("tableLogout")	
	' si l'utilisateur n'est pas connecté, on préremplie son login et mdp et on le connecte automatiquement
	if LogOut.getattribute("class") = "display-none"  And  motdepasse <> "" then
			Set LeLogin = IEDoc.getElementById("login")
			Set Password = IEDoc.getElementById("password")			
			LeLogin.Value = utilisateur
			Password.Value = motdepasse
			
			Call IEDoc.parentWindow.execScript("loginUser()","JavaScript")			
	end if
	'InputTexte.Value = Replace(Replace(Application.ActiveWindow.title.Selection(),"$a",""),"$b",", ")

    Call IEDoc.parentWindow.execScript("valueRech('Init','true')","JavaScript")
	
	filtre = false
	zoneSelection = Left(ligneSelection,3)
	if len(zoneSelection) > 2 then		  
			
		zoneD2 = valeurSousZone(zoneSelection, "$2", ligneSelection)
		index1 = "" 	
		
		Select Case zoneSelection
			Case "500", "605"
				index1 = "Titre"
				
				envoieValeurIdRef "z230_a", zoneSelection, "$a", ligneSelection
				
			Case "600", "700", "701", "702"
				index1 = "Nom de personne"		
				
				if (zoneSelection="600" and zoneD2="rameau") or (zoneSelection="700" or zoneSelection="701" or zoneSelection="702") then
					envoieValeurIdRef "z200_a", zoneSelection, "$a", ligneSelection
					zoneG = valeurSousZone(zoneSelection, "$g", ligneSelection)
					if len(zoneG)>1 then
						envoieValeurIdRef "z200_b", zoneSelection, "$g", ligneSelection
					else
						envoieValeurIdRef "z200_b", zoneSelection, "$b", ligneSelection
					end if
					envoieValeurIdRef "z200_f", zoneSelection, "$f", ligneSelection
					envoieValeurIdRef "z200_c", zoneSelection, "$c", ligneSelection
				end if
				
			Case "601", "710", "711", "712"
				index1 = "Nom de collectivit\xE9"
				ind1 = valeurIndice(zoneSelection, 0, ligneSelection)
				if ind1="0" then
					filtre = true
					Call IEDoc.parentWindow.execScript("valueRech('Index1','"+index1+"')","JavaScript")
					Call IEDoc.parentWindow.execScript("valueRech('Filtre1','Type de notice/Collectivit\xE9')","JavaScript")
				elseif ind1="1" then
					index1 = "Congrès"
				end if
				
				if (zoneSelection="601" and zoneD2="rameau") or (zoneSelection="710" or zoneSelection="711" or zoneSelection="712") then
					envoieValeurIdRef "z210_a", zoneSelection, "$a", ligneSelection
					envoieValeurIdRef "z210_b", zoneSelection, "$b", ligneSelection
					'$c repetable mais bug côté IdRef
					envoieValeurIdRef "z210_c", zoneSelection, "$c", ligneSelection
					'envoieDesValeursIdRef "z210_c", zoneSelection, "$c", ligneSelection
				end if
				
			Case "602", "720", "721", "722"
				index1 = "Famille"
				
				if (zoneSelection="602" and zoneD2="rameau") or (zoneSelection="720" or zoneSelection="721" or zoneSelection="722") then
					envoieValeurIdRef "z220_a", zoneSelection, "$a", ligneSelection
					envoieValeurIdRef "z220_c", zoneSelection, "$c", ligneSelection
					envoieValeurIdRef "z220_d", zoneSelection, "$d", ligneSelection
				end if
				
			Case "604"
				index1 = "Auteur-Titre"
				
				envoieValeurIdRef "z240_a", zoneSelection, "$a", ligneSelection
				envoieValeurIdRef "z240_t", zoneSelection, "$t", ligneSelection
				
			Case "606"
				index1 = "Nom commun"
				if zoneD2="rameau" or zoneD2="fmesh" then
					filtre = true
					Call IEDoc.parentWindow.execScript("valueRech('Index1','"+index1+"')","JavaScript")
					if zoneD2="rameau" then
						Call IEDoc.parentWindow.execScript("valueRech('Filtre1','Type de notice/Rameau')","JavaScript")
					elseif zoneD2="fmesh" then
						Call IEDoc.parentWindow.execScript("valueRech('Filtre1','Type de notice/Fmesh')","JavaScript")
					end if
				end if
				
			Case "607"
				index1 = "Nom g\xE9ographique"
				
				envoieValeurIdRef "z215_a", zoneSelection, "$a", ligneSelection
				
			Case "608"
				index1 = "Forme ou genre Rameau"
				
			Case "616"
				index1 = "Nom de marque"
				
			Case Else
				index1 = "Nom de personne"
				
		End Select
		
		zone328b = valeurSousZone("328","$b","")
		if len(zone328b)>0 then
			zone328c = valeurSousZone("328","$c","")
			zone328d = valeurSousZone("328","$d","")
			zone328e = valeurSousZone("328","$e","")
			
			envoi340 = ""
			
			zoneD4 = valeurSousZone(zoneSelection, "$4", ligneSelection)
			Select Case zoneD4
				Case "070"
					envoi340 = "Auteur"
				Case "727"
					envoi340 = "Directeur"
				Case "555"
					envoi340 = "Membre du jury"
				Case "956"
					envoi340 = "Président du jury"
				Case "727"
					envoi340 = "Directeur"
				Case "958"
					envoi340 = "Rapporteur"
				Case Else
					envoi340 = ""
			End Select
			
			if (len(envoi340)>0) then
				if (Instr(zone328b,"Mémoire")>0) then
					envoi340 = envoi340 + " d'un " + zone328b
				else
					envoi340 = envoi340 + " d'une " + zone328b
				end if
				
				if (len(zone328c)>0) then
					envoi340 = envoi340 + " en " + zone328c
				end if
				
				if (len(zone328e)>0) then
					envoi340 = envoi340 + " à " + zone328e
				end if
				
				if (len(zone328d)>9) then
					envoi340 = envoi340 + " en " + Mid(zone328d,7,4)
				end if
				
				envoieValIdRef "z340_a", envoi340
			end if
		end if
		
		zone200a = valeurSousZone("200","$a","")
		if len(zone200a)>1 then
			zone200d = valeurSousZone("200","$d","")
			zone200e = valeurSousZone("200","$e","")
			zone200f = valeurSousZone("200","$f","")
			zone200g = valeurSousZone("200","$g","")
			zone200h = valeurSousZone("200","$h","")
			zone210a = valeurSousZone("210","$a","")
			zone210c = valeurSousZone("210","$c","")
			zone210d = valeurSousZone("210","$d","")
			
			envoi810 = Replace(zone200a,"@","") + zone200h
			
			if (len(zone200d)>0) then
				envoi810 = envoi810 + " = " + zone200d
			end if
			
			if (len(zone200e)>0) then
				envoi810 = envoi810 + " : " + zone200e
			end if
			
			if (len(zone200f)>0) then
				envoi810 = envoi810 + " / " + zone200f
			end if
			
			if (len(zone200g)>0) then
				envoi810 = envoi810 + " ; " + zone200g
			end if
			
			if (len(zone210a)>0) then
				envoi810 = envoi810 + " / " + zone210a
				
				if (len(zone210c)>0) then
					envoi810 = envoi810 + " : " + zone210c
				end if
				
				if (len(zone210d)>0) then
					envoi810 = envoi810 + " , " + zone210d + "."
				end if
			end if
			
			envoieValIdRef "z810_a", envoi810
		end if
		
		if not(filtre) then 
			Call IEDoc.parentWindow.execScript("valueRech('Index1','"+index1+"')","JavaScript")
		end if
		
		'Valeur envoyée en recherche dans l'encart de recherche d'IdRef
		valeurEnvoyee = ""
		zoneDa = valeurSousZone(zoneSelection, "$a", ligneSelection)	
		if (len(zoneDa)>0) then
			valeurEnvoyee = zoneDa
			zoneDb = valeurSousZone(zoneSelection, "$b", ligneSelection)
			if (len(zoneDb)>0) then
				valeurEnvoyee = valeurEnvoyee&" "&zoneDb
			end if
		else
			'Cas speciaux des 606 et 607 : si pas de $a on essaie de prendre la valeur d'une subdivision
			if zoneSelection="606" or zoneSelection="607" then
				zoneDx = valeurSousZone(zoneSelection, "$x", ligneSelection)
				if (len(zoneDx)>0) then
					valeurEnvoyee = zoneDx
				else
					zoneDy = valeurSousZone(zoneSelection, "$y", ligneSelection)
					if (len(zoneDy)>0) then
						valeurEnvoyee = zoneDy
					else
						zoneDz = valeurSousZone(zoneSelection, "$z", ligneSelection)
						if (len(zoneDz)>0) then
							valeurEnvoyee = zoneDz
						end if
					end if
				end if
			end if
		end if
		Call IEDoc.parentWindow.execScript("valueRech('Index1Value','"+valeurEnvoyee+"')","JavaScript")
		
	end if
	
	' rempli les champs Idref avec les champs de la notice biblio pour la création d'une nouvelle autorité
	envoieValeurIdRef "z101_a", "101", "$a", ""
	envoieValeurIdRef "z102_a", "102", "$a", ""
	envoieValeurIdRef "z103_a", "103", "$a", ""

	'Lance automatiquement la recherche (si pas de filtre selectionne, sinon bug d'IdRef (filtre non selectionné)
	'Bug si pas de résultat : retour à la page de recherche et tous les filtres affichés
	if not(filtre) then
		Call IEDoc.parentWindow.execScript("valueRech('AutoClick','true')","JavaScript")
		Call IEDoc.parentWindow.execScript("valueRech('End','true')","JavaScript")	
		Call IEDoc.parentWindow.execScript("lanceRech()","JavaScript")
	end if

	 'remet le focus sur la zone saisie
	 ligneSelection = application.ActiveWindow.title.findTag(ligneSelection, 0, false, true, true)
     
	Set IE.document.all("Lier").onclick = GetRef("rapatrie")
	'IE.document.focus()
	IE.Visible = False
	IE.Visible = True	         
	'Application.windows.item(0).minimize
	'Application.Visible = True
	'Application.Visible = False
	 
end if
	
End Sub


Sub rapatrie()

	application.ActiveWindow.title.EndOfField()
	application.ActiveWindow.title.StartOfField(True)  
	ligneSelection = application.ActiveWindow.title.selection()
	
	For Each aElement In IE.document.all
      
		if aElement.ClassName = "detail_ppn2" Then
			For Each bElement In aElement.all
				If bElement.ClassName = "detail_value" Then
					'Par défaut la valeur sera remplacée par les indicateurs, suivis du PPN trouvé dans IdRef
					valeur = "$3"&bElement.innerText					
					indicateurs = Mid(ligneSelection,5,2)
					if instr(indicateurs,"$")>0 then
						indicateurs = ""
					end if
							
					zoneD2 = valeurSousZone(zoneSelection, "$2", ligneSelection)
					zoneD4 = valeurSousZone(zoneSelection, "$4", ligneSelection)
					'Cas speciaux des 606 et 607 avec subdivisions possibles
					zoneSelection = left(application.ActiveWindow.title.TagAndSelection(),3)						
					
					if zoneSelection="606" or zoneSelection="607" Then
						zoneA = valeurSousZone(zoneSelection, "$a", ligneSelection)
						if len(zoneA)>0 then
							valeur = Replace(ligneSelection, "$a"&zoneA, "$3"&bElement.innerText)
						else
							zoneX = valeurSousZone(zoneSelection, "$x", ligneSelection)
							if len(zoneX)>0 then
								valeur = Replace(ligneSelection, "$x"&zoneX, "$3"&bElement.innerText)
							else
								zoneY = valeurSousZone(zoneSelection, "$y", ligneSelection)
								if len(zoneY)>0 then
									valeur = Replace(ligneSelection, "$y"&zoneY, "$3"&bElement.innerText)
								else
									zoneZ = valeurSousZone(zoneSelection, "$z", ligneSelection)
									if len(zoneZ)>0 then
										valeur = Replace(ligneSelection, "$z"&zoneZ, "$3"&bElement.innerText)
									end if
								end if
							end if
						end if
						valeur = Replace(valeur,"##", "")
						valeur = Replace(valeur,zoneSelection&" ", "")
					else 
						'On conserve les $2 et $4 dans les autres cas
						if len(zoneD2)>0 then
							valeur = valeur & "$2" & zoneD2 
						elseif len(zoneD4)>0 then
							valeur = valeur & "$4" & zoneD4 
						end if
					end if
					
					ligneSelection = application.ActiveWindow.title.findTag(ligneSelection, 0, false, true, true)
					Application.ActiveWindow.title.insertText(indicateurs&valeur)
									
				End If
			Next
		End If
	Next
	
	IE.Quit
	Set IE = Nothing 
	'Application.ActiveWindow.caption = Entry
End Sub	

Function valeurIndice(zone, indice, laLigne) 
	ligne = laLigne
	valeur = ""
	application.ActiveWindow.title.StartOfBuffer (false)
	if ligne = "" then 
		ligne = application.ActiveWindow.title.findTag(zone, 0, false, true, true)
	end if
	if len(ligne) > 4 then
		if indice = 0 then
			valeur = Mid(ligne,5,1)
		elseIf indice = 1 then
			valeur = Mid(ligne,6,1)
		end if
	end if	
	valeurIndice=valeur
End Function

Function valeurSousZone(zone, dollar, laLigne) 
	ligne = laLigne
	valeur = ""
	application.ActiveWindow.title.StartOfBuffer (false)
	if ligne = "" then 
		'MsgBox zone+" "+dollar+" |"+ligne+"|"
		ligne = application.ActiveWindow.title.findTag(zone, 0, false, true, true)
	end if
	debutsouszone = instr(ligne,dollar)
	if debutsouszone <> 0 then
		ligne = mid(ligne,debutsouszone+2)
		finsouszone = instr(ligne,"$")
		if finsouszone = 0 then
			finsouszone = len(ligne) + 1
		end if
		valeur = left(ligne,finsouszone-1)
		valeur = Replace(valeur,"@","")
	end if	
	valeurSousZone=valeur
End Function

Sub envoieValeurIdRef(zoneIdRef, zone, dollar, ligne)
	valeur = valeurSousZone(zone, dollar, ligne)
	envoieValIdRef zoneIdRef, valeur
End Sub

Sub envoieValIdRef(zoneIdRef, valeur)
	if len(valeur) >0 then
		valeur = Replace(valeur,"'","\'")
		Call IEDoc.parentWindow.execScript("valueRech('"+zoneIdRef+"','"+ valeur +"')","JavaScript")
	end if
End Sub

Sub envoieDesValeursIdRef(zoneIdRef, zone, dollar, laLigne)
	ligne = laLigne
	valeur = ""
	application.ActiveWindow.title.StartOfBuffer (false)
	if ligne = "" then 
		ligne = application.ActiveWindow.title.findTag(zone, 0, false, true, true)
	end if
	
	cnt = 1
	debutsouszone = instr(ligne,dollar)
	while debutsouszone <> 0 and len(ligne)>0
		ligne = mid(ligne,debutsouszone+2)
		finsouszone = instr(ligne,"$")
		if finsouszone = 0 then
			finsouszone = len(ligne) + 1
		end if
		valeur = left(ligne,finsouszone-1)
		if len(valeur) >0 then
			Call IEDoc.parentWindow.execScript("valueRech('"+zoneIdRef+"_"&cnt&"','"+ valeur +"')","JavaScript")
		end if
		debutsouszone=finsouszone
		cnt = cnt + 1
	Wend
	
End Sub



