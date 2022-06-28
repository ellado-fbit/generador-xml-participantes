import { read, utils } from 'xlsx'
import moment from 'moment'

export const generateXml = (content) => {
  try {
    const excelFile = read(content, { cellDates: true })
    const participantes = utils.sheet_to_json(excelFile.Sheets['Bonificados'])
    if (participantes.length === 0) throw new Error(`La página 'Bonificados' no existe o está vacía.`)

    // sacamos grupos
    const grupos = participantes.map(x => x['Grupo'])
    grupos.forEach(gr => {
      if (!gr) throw new Error(`Existen filas con la columna Grupo no definido.`)
    })
    const uniqueGroups = grupos.filter((c, index) => grupos.indexOf(c) === index)

    const participantesPorGrupos = {}
    uniqueGroups.forEach(gr => {
      participantesPorGrupos[gr] = participantes.filter(par => par['Grupo'] === gr)
    })

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

    xml = xml + '<grupos>\n'

    uniqueGroups.forEach(grupo => {

      xml = xml + `<grupo>\n`
      xml = xml + `   <idAccion>${participantesPorGrupos[grupo][0]['Acción']}</idAccion>\n`
      xml = xml + `   <idGrupo>${grupo}</idGrupo>\n`
      xml = xml + '   <participantes>\n'

      participantesPorGrupos[grupo].forEach((participante, index) => {
        xml = xml + '      <participante>\n'

        // nif
        // ----
        if (!participante['Número NIF/NIE']) {
          throw new Error(`No existe la columna 'Número NIF/NIE' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <nif>${participante['Número NIF/NIE'].trim().toUpperCase()}</nif>\n`
        }

        // N_TIPO_DOCUMENTO
        // -----------------
        if (!participante['Tipo Documento']) {
          throw new Error(`No existe la columna 'Tipo Documento' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          if (!['NIF', 'NIE'].includes(participante['Tipo Documento'])) {
            throw new Error(`En la fila ${index + 1} del grupo ${grupo} el campo 'Tipo Documento' no es válido, valores permitidos: 'NIF' o 'NIE'.`)
          } else {
            const tipos = { 'NIF': 10, 'NIE': 60 }
            xml = xml + `         <N_TIPO_DOCUMENTO>${tipos[participante['Tipo Documento']]}</N_TIPO_DOCUMENTO>\n`
          }
        }

        // nombre
        // ---------
        if (!participante['Nombre']) {
          throw new Error(`No existe la columna 'Nombre' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <nombre>${participante['Nombre'].trim()}</nombre>\n`
        }

        // primerApellido
        // ---------------
        if (!participante['Apellido 1']) {
          throw new Error(`No existe la columna 'Apellido 1' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <primerApellido>${participante['Apellido 1'].trim()}</primerApellido>\n`
        }

        // segundoApellido
        // ----------------
        if (!participante['Apellido 2']) {
          throw new Error(`No existe la columna 'Apellido 2' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <segundoApellido>${participante['Apellido 2'].trim()}</segundoApellido>\n`
        }

        // niss
        // -----
        if (!participante['NISS']) {
          throw new Error(`No existe la columna 'NISS' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <niss>${participante['NISS']}</niss>\n`
        }

        // cifEmpresa
        // -----------
        if (!participante['CIF']) {
          throw new Error(`No existe la columna 'CIF' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <cifEmpresa>${participante['CIF']}</cifEmpresa>\n`
        }

        // ctaCotizacion
        // --------------
        if (!participante['Cuenta de cotización']) {
          throw new Error(`No existe la columna 'Cuenta de cotización' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <ctaCotizacion>${participante['Cuenta de cotización']}</ctaCotizacion>\n`
        }

        // fechaNacimiento
        // ----------------
        if (!participante['Fecha nacimiento']) {
          throw new Error(`No existe la columna 'Fecha nacimiento' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          let fecha = participante['Fecha nacimiento'].toISOString().split('T')[0]
          fecha = moment(fecha, 'YYYY-MM-DD').add(1, 'days').format('DD/MM/YYYY')  // add one day to solve date conversion problem
          xml = xml + `         <fechaNacimiento>${fecha}</fechaNacimiento>\n`
        }

        // email
        // ------
        if (!participante['Email']) {
          throw new Error(`No existe la columna 'Email' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <email>${participante['Email']}</email>\n`
        }

        // telefono
        // ---------
        if (!participante['Teléfono']) {
          throw new Error(`No existe la columna 'Teléfono' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <telefono>${participante['Teléfono']}</telefono>\n`
        }

        // sexo
        // -----
        if (!participante['Sexo']) {
          throw new Error(`No existe la columna 'Sexo' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          if (!['Mujer', 'Hombre'].includes(participante['Sexo'])) {
            throw new Error(`En la fila ${index + 1} del grupo ${grupo} el campo 'Sexo' no es válido, valores permitidos: 'Mujer' o 'Hombre'.`)
          } else {
            const tipos = { 'Mujer': 'F', 'Hombre': 'M' }
            xml = xml + `         <sexo>${tipos[participante['Sexo']]}</sexo>\n`
          }
        }

        // categoriaprofesional
        // ---------------------
        if (!participante['Categoría profesional']) {
          throw new Error(`No existe la columna 'Categoría profesional' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <categoriaprofesional>${participante['Categoría profesional'].trim().split('. ')[0]}</categoriaprofesional>\n`
        }

        // grupocotizacion
        // ----------------
        if (!participante['Grupo de cotización']) {
          throw new Error(`No existe la columna 'Grupo de cotización' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <grupocotizacion>${participante['Grupo de cotización'].trim().split('. ')[0]}</grupocotizacion>\n`
        }

        // nivelestudios
        // --------------
        if (!participante['Nivel de estudios']) {
          throw new Error(`No existe la columna 'Nivel de estudios' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          xml = xml + `         <nivelestudios>${participante['Nivel de estudios'].trim().split('. ')[0]}</nivelestudios>\n`
        }

        // DiplomaAcreditativo
        // --------------------
        if (!participante['Diploma acreditativo']) {
          throw new Error(`No existe la columna 'Diploma acreditativo' en la fila ${index + 1} del grupo ${grupo}`)
        } else {
          if (!['Sí', 'No'].includes(participante['Diploma acreditativo'])) {
            throw new Error(`En la fila ${index + 1} del grupo ${grupo} el campo 'Diploma acreditativo' no es válido, valores permitidos: 'Sí' o 'No'.`)
          } else {
            const tipos = { 'Sí': 'S', 'No': 'N' }
            xml = xml + `         <DiplomaAcreditativo>${tipos[participante['Diploma acreditativo']]}</DiplomaAcreditativo>\n`
          }
        }

        xml = xml + '       </participante>\n'
      })

      xml = xml + '   </participantes>\n'
      xml = xml + '</grupo>\n'

    })

    xml = xml + '</grupos>\n'

    return xml

  } catch (error) {
    return error
  }
}