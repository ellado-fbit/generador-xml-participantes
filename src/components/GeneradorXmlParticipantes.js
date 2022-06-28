import { Input } from 'antd'
import { useState } from 'react'
import { generateXml } from './utils/generateXml'

const { TextArea } = Input

export const GeneradorXmlParticipantes = () => {
  const [xmlContent, setXmlContent] = useState('')

  const captureFile = (event) => {
    setXmlContent('')
    const file = event.target.files[0]
    event.target.value = null
    const fileReader = new FileReader()

    if (file) {
      fileReader.onloadend = () => {
        let xml
        try {
          xml = generateXml(fileReader.result)
          setXmlContent(xml)
        } catch (error) {
          xml = error.message
        }
      }
      fileReader.readAsArrayBuffer(file)
    }
  }

  return (
    <>
      <input
        type="file"
        accept={'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
        onChange={captureFile}
        onClick={() => { setXmlContent('') }}
      />
      <div>
        <br />
        <TextArea rows={20} value={xmlContent} />
      </div>
    </>
  )
}
