# reportOrariViaggiatreno
Il contenuto di questo repository è il frutto di una ricerca personale, non esiste alcuna affiliazione con Trenitalia o con altri organi delle Ferrovie dello Stato.

Questo codice è da considerarsi una bozza: sono stati implementati solo i metodi principali e probabilmente esistono molti casi limite che causano eccezioni non gestite. 

Lo scopo di questo codice consiste nella generazione di un file excel sui ritardi di una lista di treni censiti.

## Requisiti
* Python 3.1 
* xlsxwriter (https://xlsxwriter.readthedocs.io/)

##Documenti

Lo script utilizza fondamentalmente le API recuperabili direttamente dal portale Viaggiatreno mediante esplorazione delle chiamate di rete del browser.

L'API utilizzata per la costruzione del report excel è
* http://www.viaggiatreno.it/infomobilita/resteasy/viaggiatreno/andamentoTreno/[codice_stazione_]/[codice_treno]/[timestamp_DataOdierna]
e prevede tre parametri in ingresso definiti dal codice della stazione, numero del treno e la data odierna convertita in timestamp utc.

La coppia codice_stazione e codice_treno è recuperabile dalla seguente API:
* http://www.viaggiatreno.it/infomobilita/resteasy/viaggiatreno/cercaNumeroTrenoTrenoAutocomplete/[numero_treno]

All'interno del codice è definita una HashMap (chiave - valore) necessaria per definire la lista dei treni che si vogliono tracciare. 
All'interno del foglio excel verrà creato un tab per ciascun treno e per ciascun treno verranno forniti i ritardi giornalieri suddivisi per le varie stazioni.

##Note
Il servizio é svolto a scopo di studio, le informazioni non sono pubblicamente accessibili nonostante l'azienda sia pubblica.

