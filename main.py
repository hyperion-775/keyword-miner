#Written by hyperion-775

#Imports
import pandas as pd
import numpy as np

#Insert complaint file 1 here
path1 = r''
#Insert complaint file 2 here
path2 = r''
#Insert SOE file here
path3 = r''

#In complaint file 1, insert the column letter of the ticket number inside the quotation marks
ticketnumbercolumn1 = ""
#In complaint file 2, insert the column letter of the ticket number inside the quotation marks
ticketnumbercolumn2 = ""
#In complaint file 1, insert the column letter of the ticket notes inside the quotation marks
ticketnotescolumn1 = ""
#In complaint file 2, insert the column letter of the ticket notes inside the quotation marks
ticketnotescolumn2 = ""
#In complaint file 1, insert the column letter of the conclusion summary inside the quotation marks
conclusionsummarycolumn1 = ""
#In complaint file 2, insert the column letter of the conclusion summary inside the quotation marks
conclusionsummarycolumn2 = ""
#In the SOE file, insert the column letter of the SOE ID inside the quotation marks
SOEIDcolumn = ""
#In the SOE file, insert the column letter of the SOE keywords inside the quotation marks
SOEKeywordcolumn = ""

#Insert output file name here
outputfilename = ''

def dfToList(path, column):
    dfkeywords = pd.read_excel(path, usecols= column)
    dfkeywordsList = dfkeywords.values.tolist()
    return dfkeywordsList

ticketnumber1 = dfToList(path1, ticketnumbercolumn1)
ticketnotes1 = dfToList(path1, ticketnotescolumn1)
conclusionsummary1 = dfToList(path1, conclusionsummarycolumn1)

ticketnumber2 = dfToList(path2, ticketnumbercolumn2)
ticketnotes2 = dfToList(path2, ticketnotescolumn2)
conclusionsummary2 = dfToList(path2, conclusionsummarycolumn2)

ticketnumber = np.ravel([*ticketnumber1, *ticketnumber2])
ticketnotesnativelower = np.ravel([*ticketnotes1, *ticketnotes2])
concsumlower = np.ravel([*conclusionsummary1, *conclusionsummary2])

concsum = [x.upper() for x in concsumlower]
ticketnotesnative = [tnotesn.upper() for tnotesn in ticketnotesnativelower]
listOfListSOEID = dfToList(path3, SOEIDcolumn) 
listofSOEID = np.ravel(listOfListSOEID)
listOfListSOEKeywordsinitial = np.ravel(dfToList(path3, SOEKeywordcolumn))

listOfListSOEKeywordscaps = [caps.upper() for caps in listOfListSOEKeywordsinitial]
listOfListSOEKeywords = [seperate.split(', ') for seperate in listOfListSOEKeywordscaps]

listOfIdList = [[]]*len(listOfListSOEKeywords)
listOfTnotesN = [[]]*len(listOfListSOEKeywords)
listOfConcSum = [[]]*len(listOfListSOEKeywords)
listOfDataCombined = [[]]*len(listOfListSOEKeywords)

hitcounter = 0
for i in listOfListSOEKeywords: 
    keywordscounter = listOfListSOEKeywords.index(i)
    listOfIdList[keywordscounter]=[]
    listOfDataCombined[keywordscounter]=[]

    hitcounter = 0
    for j in i:
                tempListID = []
                tempListConcSum = []
                tempListTicketNotesN = []
                tempListTicketNotesT = []
                tempListTicketDataCombined = []

                    
                                           
                for k in concsum:    
                        conclusionsummarycounter = concsum.index(k)
                        if " AND " in j:
                            andsplit = j.split(" AND ")
                            lentracker = 0
                            for anditeration in andsplit:
                                if "DOES NOT CONTAIN " in anditeration:
                                    dncsplit = anditeration.split("DOES NOT CONTAIN ").pop()
                                    if " (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)" in dncsplit:
                                        if anditeration.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in k and anditeration.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in (ticketnotesnative[conclusionsummarycounter]):
                                            lentracker = lentracker + 1
                                    elif dncsplit not in k:
                                        lentracker = lentracker + 1
                                elif " (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)" in anditeration:
                                    if anditeration.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] in k and anditeration.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] in (ticketnotesnative[conclusionsummarycounter]):#
                                        lentracker = lentracker + 1
                                elif anditeration in k:
                                    lentracker = lentracker + 1
                            if lentracker == len(andsplit):

                                tempListID.append(conclusionsummarycounter)
                                tempListConcSum.append(concsum[conclusionsummarycounter])
                                tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))

                                hitcounter = hitcounter + 1
                                lentracker = 0
                            else:
                                lentracker = 0
                        elif "DOES NOT CONTAIN " in j:
                                if " (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)" in j.split("DOES NOT CONTAIN ").pop():
                                    if j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in k and j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in (ticketnotesnative[conclusionsummarycounter]):#
                                        tempListID.append(conclusionsummarycounter)
                                        tempListConcSum.append(concsum[conclusionsummarycounter])
                                        tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                        tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))                           
                                        hitcounter = hitcounter + 1
                                elif j.split("DOES NOT CONTAIN ").pop() not in k:
                                    tempListID.append(conclusionsummarycounter)
                                    tempListConcSum.append(concsum[conclusionsummarycounter])
                                    tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                    tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))
                                    hitcounter = hitcounter + 1 
                        elif " (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)" in j:
                            if j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] in k and j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] in (ticketnotesnative[conclusionsummarycounter]):#
                                tempListID.append(conclusionsummarycounter)
                                tempListConcSum.append(concsum[conclusionsummarycounter])
                                tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))


                                hitcounter = hitcounter + 1
                        elif j in k:
                            tempListID.append(conclusionsummarycounter)
                            tempListConcSum.append(concsum[conclusionsummarycounter])
                            tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                            tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))

                            hitcounter = hitcounter + 1
                                                
                        if len(tempListID) > 10:
                            tempListID = []
                            for item in tempListID:
                                if " AND " in j:
                                    andsplit = j.split(" AND ")
                                    # print(andsplit)
                                    lentracker = 0
                                    for anditeration in andsplit:
                                        if "DOES NOT CONTAIN " in anditeration:
                                            dncsplit = anditeration.split("DOES NOT CONTAIN ").pop()
                                            if " (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)" in dncsplit:
                                                if anditeration.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in (concsum[item] and anditeration.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in (ticketnotesnative[item])):
                                                    lentracker = lentracker + 1
                                            elif dncsplit not in (k and (ticketnotesnative[conclusionsummarycounter])):
                                                lentracker = lentracker + 1
                                        elif " (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)" in anditeration:
                                            if anditeration.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] in concsum[item] and anditeration.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] in (ticketnotesnative[conclusionsummarycounter]):#
                                                lentracker = lentracker + 1
                                        elif anditeration in concsum[item] and anditeration in (ticketnotesnative[conclusionsummarycounter]):
                                            lentracker = lentracker + 1
                                    if lentracker == len(andsplit):

                                        tempListID.append(conclusionsummarycounter)
                                        tempListConcSum.append(concsum[conclusionsummarycounter])
                                        tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                        tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))


                                        #hitcounter = hitcounter + 1
                                        lentracker = 0
                                    else:
                                        lentracker = 0
                                elif "DOES NOT CONTAIN " in j:
                                        if " (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)" in j.split("DOES NOT CONTAIN ").pop():
                                            if j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in concsum[item] and j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in (ticketnotesnative[conclusionsummarycounter]):#
                                                tempListID.append(conclusionsummarycounter)
                                                tempListConcSum.append(concsum[conclusionsummarycounter])
                                                tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                                tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))
                                                hitcounter = hitcounter + 1
                                        elif j.split("DOES NOT CONTAIN ").pop() not in concsum[item] and j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] not in (ticketnotesnative[conclusionsummarycounter]):
                                            tempListID.append(conclusionsummarycounter)
                                            tempListConcSum.append(concsum[conclusionsummarycounter])
                                            tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                            tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))

                                elif " (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)" in j:
                                    if j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] in concsum[item] and j.split(' (IN BOTH TICKET NOTES ALONG WITH CONCLUSION SUMMARY COLUMNS)')[0] in (ticketnotesnativelower[conclusionsummarycounter]):#
                                        tempListID.append(conclusionsummarycounter)
                                        tempListConcSum.append(concsum[conclusionsummarycounter])
                                        tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                        tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))
                
                                        #hitcounter = hitcounter + 1
                                elif ((j in concsum[item]) and (j in (ticketnotesnative[conclusionsummarycounter]))):
                                        tempListID.append(conclusionsummarycounter)
                                        tempListConcSum.append(concsum[conclusionsummarycounter])
                                        tempListTicketNotesN.append(ticketnotesnative[conclusionsummarycounter])
                                        tempListTicketDataCombined.append('TICKET NUMBER: '+ str(ticketnumber[conclusionsummarycounter]) + '\n\nCONCLUSION SUMMARY: ' + str(concsumlower[conclusionsummarycounter]) +  '\n\nTICKET NOTES: ' + str(ticketnotesnativelower[conclusionsummarycounter]))
        
                                    #hitcounter = hitcounter + 1 

                                    #hitcounter = hitcounter + 1                  

                listOfIdList[keywordscounter] = [*listOfIdList[keywordscounter], *tempListID]
                listOfTnotesN[keywordscounter] = [*listOfTnotesN[keywordscounter], *tempListTicketNotesN]
                listOfConcSum[keywordscounter] = [*listOfConcSum[keywordscounter], *tempListConcSum]
                listOfDataCombined[keywordscounter] = listOfDataCombined[keywordscounter] + tempListTicketDataCombined
                finallistofIdlist = [ticketnumber[i] for i in listOfIdList]
                # finalDataCombined = 
                hitcounter=0
#                print("Hitcounter Round 2 ", hitcounter, len(tempListID), j)   
#                print(tempListID)

finallistofIdliststring = [str(i) for i in finallistofIdlist]                  
listOfTnotesNString = [str(i) for i in listOfTnotesN]                  
listOfConcSumString = [str(i) for i in listOfConcSum]
listOfListSOEKeywordsstring = [str(i) for i in listOfListSOEKeywords]
print(pd.DataFrame(listOfDataCombined))

#print(listOfIdListString)
#outputdf = pd.DataFrame(listOfIdListString)

#outputdf = pd.DataFrame({'SOEID': listofSOEID, 'Ticket Numbers': listOfIdListString, 'Ticket Notes Translated': listOfTnotesTString,'Ticket Notes Native': listOfTnotesNString, 'Conclusion Summary': listOfConcSumString})
#outputdf1 = pd.DataFrame({listofSOEID, listOfIdListString, listOfDataCombined})
outputdf1 = [pd.DataFrame(listofSOEID).transpose(), pd.DataFrame(listOfListSOEKeywordsstring).transpose(), pd.DataFrame(finallistofIdliststring).transpose(), pd.DataFrame(listOfDataCombined).transpose()]
# outputdf2 = pd.DataFrame(listOfDataCombined).transpose()

outputdf2 = pd.concat(outputdf1)
outputdf2.to_excel(outputfilename, sheet_name='Sheet1', index=False)
