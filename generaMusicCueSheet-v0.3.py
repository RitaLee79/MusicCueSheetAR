# NOTAS:
# Limpieza de proyecto: nombre "fcp mcs Agosto f1" / en el nivel principal solo carpetas con dia del mes "01 02 03 .."
# en las carpetas solo secuencias no precomps,..
#
# OJO con las secuencias que se quedan restos de canciones al final pues la deteccion no esta implementada
#
# Los fx de sonido que se vayan colando anhardirlos a blackListFX siguiendo el patron

import xml.etree.ElementTree as ET
import pyexcel

import tkFileDialog

fileToOpen = tkFileDialog.askopenfilename(initialdir = "~/Desktop",
                                          title = "Elige el XML del proyecto",
                                          filetypes = (("Archivos XML","*.xml"),("Todos los archivos","*.*")))

fileToWrite = tkFileDialog.asksaveasfilename(title = "Elige el archivo de destino", defaultextension='xls') #'Parte Musicas FinalCut.xls'

blackListFX = ['Impact Hit Reverb','Delay woosh','Flame woosh','Servo 2',
               'Windy swish','FX RAF Rebotes varios','Fast swish','Firey swoosh',
               'TELEMETRIA','Heavy stomp','Bassy swoosh','Impact Metal.caf','Camera 1',
               'Disparo con eco2','Disparo 2','Swoosh low']

######################################################################################
# Pasar FRAMES --> TIMECODE
#
# fusilado de https://github.com/eoyilmaz/timecode/blob/master/timecode/__init__.py
#
######################################################################################

def frames_to_tc(frames):
    """Converts frames back to timecode
    :returns str: the string representation of the current time code
    """
    ffps = 25.0

    # Number of frames in an hour
    frames_per_hour = int(round(ffps * 60 * 60))
    # Number of frames in a day - timecode rolls over after 24 hours
    frames_per_24_hours = frames_per_hour * 24
    # Number of frames per ten minutes
    frames_per_10_minutes = int(round(ffps * 60 * 10))
    # Number of frames per minute is the round of the framerate * 60 minus
    # the number of dropped frames
    frames_per_minute = int(round(ffps) * 60)

    frame_number = frames - 1

    if frame_number < 0:
        # Negative time. Add 24 hours.
        frame_number += frames_per_24_hours

    # If frame_number is greater than 24 hrs, next operation will rollover
    # clock
    frame_number %= frames_per_24_hours

    ifps = 25.0
    frs = frame_number % ifps
    secs = (frame_number // ifps) % 60
    mins = ((frame_number // ifps) // 60) % 60
    hrs = (((frame_number // ifps) // 60) // 60)

    return "%02d:%02d:%02d.%02d" % (hrs,
                                     mins,
                                     secs,
                                     frs)


############################
#
# Analizamos archivo XML
#
############################


xml_file = open(fileToOpen, "r")

tree = ET.parse(xml_file)
root = tree.getroot()

musicCueSheet = [[],[],[],[]]

for bins in root.iter('bin'):

    w_DiaDelMes = bins.find('name').text  # .encode('ascii', 'replace')                  # dia del mes

    for sequences in bins.iter('sequence'):

        isSequence = sequences.find('name')
        if isSequence is None:
            continue

        seqWithAudio = sequences.find('media')
        if seqWithAudio is None:
            continue

        w_NombreSeq = sequences.find('name').text  # .encode('ascii', 'replace')        # nombre de la seq
        seqMusicCueSheet = [[], []]

        desdeA5 = 4

        for tracks in sequences.findall('./media/audio/track'):

            #  saltar A1 - A4
            if desdeA5 < 8:
                desdeA5 += 1
                continue

            rawTrackList = []

            # Analizamos los todos clips del track
            for audioClips in tracks:
                if audioClips.tag == 'transitionitem':

                    rawTrackList.append([int(audioClips.findtext('start')), int(audioClips.findtext('end'))])

                elif audioClips.tag == 'clipitem':

                    if audioClips.findtext('enabled') == 'FALSE':
                         isEnabled = 0
                    else:
                         isEnabled = 1
                    rawTrackList.append([audioClips.findtext('name'), int(audioClips.findtext('start')),
                                         int(audioClips.findtext('end')), audioClips.find('file').find('media'),
                                         audioClips.find('sourcetrack').findtext('trackindex'), isEnabled])

            # Calculamos las canciones del track
            for clipitem in range(len(rawTrackList)):

                # Si es transicion pasamos al siguiente
                if len(rawTrackList[clipitem]) == 2:
                    continue

                # Si esta desactivado pasamos al siguiente
                if rawTrackList[clipitem][5] == 0:
                    continue

                w_TrackName = rawTrackList[clipitem][0]

                # Chequeamos si el item esta en blackListFX:
                if w_TrackName.startswith(tuple(blackListFX)):
                    continue

                # Comprobamos si es 2a pista estereo
                if rawTrackList[clipitem][3] is None and rawTrackList[clipitem][4] == '2':
                    continue

                # Calculamos la duracion
                if rawTrackList[clipitem][1] == -1:
                    clipStart = rawTrackList[clipitem-1][0]
                else:
                    clipStart = rawTrackList[clipitem][1]

                if rawTrackList[clipitem][2] == -1:
                    if (clipitem+2) >= len(rawTrackList):
                        clipEnd = rawTrackList[clipitem + 1][1]
                    elif len(rawTrackList[clipitem+1]) == 2 and rawTrackList[clipitem+2][1] == -1 \
                            and rawTrackList[clipitem+2][0] == rawTrackList[clipitem][0]:
                        clipEnd = rawTrackList[clipitem + 1][0]
                    else:
                        clipEnd = rawTrackList[clipitem + 1][1]
                else:
                    clipEnd = rawTrackList[clipitem][2]

                w_Duration = clipEnd - clipStart

                # Almacenamos en seqMusicCueSheet
                try:
                    yaAnhadidaIndex = seqMusicCueSheet[0].index(w_TrackName)
                except ValueError:
                    #print "no esta todavia en la lista"
                    seqMusicCueSheet[0].append(w_TrackName)
                    seqMusicCueSheet[1].append(w_Duration)
                else:
                    #print "si esta en la lista"
                    seqMusicCueSheet[1][yaAnhadidaIndex] += w_Duration

            desdeA5 += 1

        # Almacenamos en musicCueSheet las canciones de la secuencia
        for canciones in range(len(seqMusicCueSheet[0])):
            musicCueSheet[0].append(w_DiaDelMes)
            musicCueSheet[1].append(w_NombreSeq.replace(',', ' '))
            musicCueSheet[2].append(seqMusicCueSheet[0][canciones].replace(',', ' '))
            musicCueSheet[3].append(seqMusicCueSheet[1][canciones])

# Preparamos lista para exportar a xls
xlsListExport = []

for i in range(len(musicCueSheet[0])):

    xlsListExport.append([musicCueSheet[0][i], musicCueSheet[1][i],
                          musicCueSheet[2][i], frames_to_tc(musicCueSheet[3][i])])

# Exportamos archivo XLS
exportedFile = open(fileToWrite, 'w')
pyexcel.save_as(array=xlsListExport, dest_file_name=fileToWrite)
exportedFile.close()


