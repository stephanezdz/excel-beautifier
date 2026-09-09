import { useState, useEffect } from 'react'
import { useDropzone } from 'react-dropzone'
import axios from 'axios'
import * as XLSX from 'xlsx'

const THEMES = {
  professional: {
    name: 'Professionnel',
    color: '#2E74B5',
    row: '#F2F2F2'
  },
  modern: {
    name: 'Moderne',
    color: '#1E3A5F',
    row: '#F8F9FA'
  },
  clean: {
    name: 'Épuré',
    color: '#4A5568',
    row: '#FFFFFF'
  }
}

function App() {
  const [file, setFile] = useState(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const [downloadUrl, setDownloadUrl] = useState(null)
  const [error, setError] = useState(null)
  
  const [settings, setSettings] = useState({
    theme: 'professional',
    header_font_size: 14,
    header_font_weight: 'bold',
    header_color: '#2E74B5',
    alternate_row_color: '#F2F2F2',
    border_color: '#CCCCCC',
    header_alignment: 'center'
  })

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
    onDrop: (acceptedFiles) => {
      const selectedFile = acceptedFiles[0]
      setFile(selectedFile)
      setError(null)
      setDownloadUrl(null)
    }
  })

  const handleBeautify = async () => {
    if (!file) return
    
    setIsProcessing(true)
    setError(null)
    
    const formData = new FormData()
    formData.append('file', file)
    formData.append('settings', new Blob([JSON.stringify(settings)], { type: 'application/json' }))
    
    try {
      const response = await axios.post('http://localhost:8000/api/beautify', formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        },
        responseType: 'blob'
      })
      
      const url = URL.createObjectURL(response.data)
      setDownloadUrl(url)
    } catch (err) {
      setError(err.response?.data?.detail || 'Une erreur est survenue')
    } finally {
      setIsProcessing(false)
    }
  }

  const handleDownload = () => {
    if (downloadUrl) {
      const link = document.createElement('a')
      link.href = downloadUrl
      link.download = `beautified-${file.name}`
      link.click()
      URL.revokeObjectURL(downloadUrl)
      setDownloadUrl(null)
    }
  }

  const handlePreview = () => {
    if (file) {
      const reader = new FileReader()
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' })
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })
        
        console.log(jsonData)
      }
      reader.readAsArrayBuffer(file)
    }
  }

  return (
    <div style={{
      fontFamily: 'Arial, sans-serif',
      minHeight: '100vh',
      backgroundColor: '#f5f5f5',
      padding: '20px'
    }}>
      <div style={{ maxWidth: '800px', margin: '0 auto' }}>
        <h1 style={{ textAlign: 'center', color: '#2c3e50', marginBottom: '30px' }}>
          🎨 Excel Beautifier
        </h1>
        
        {/* Upload Section */}
        <div 
          {...getRootProps()}
          style={{
            border: '3px dashed #3498db',
            borderRadius: '10px',
            padding: '40px',
            textAlign: 'center',
            backgroundColor: isDragActive ? '#e8f6fd' : 'white',
            cursor: 'pointer',
            marginBottom: '20px'
          }}
        >
          <input {...getInputProps()} />
          <p style={{ fontSize: '18px', color: '#3498db' }}>
            {isDragActive 
              ? 'Déposez votre fichier Excel ici' 
              : 'Glissez-déposez votre fichier Excel ici ou cliquez pour parcourir'}
          </p>
          <p style={{ fontSize: '14px', color: '#7f8c8d' }}>
            Supporte les fichiers .xlsx
          </p>
        </div>

        {/* Error Display */}
        {error && (
          <div style={{
            backgroundColor: '#ffdddd',
            border: '1px solid #f8d7da',
            borderRadius: '5px',
            padding: '15px',
            marginBottom: '20px',
            color: '#721c24'
          }}>
            <strong>Erreur:</strong> {error}
          </div>
        )}

        {/* File Info */}
        {file && (
          <div style={{
            backgroundColor: '#d4edda',
            border: '1px solid #c3e6cb',
            borderRadius: '5px',
            padding: '15px',
            marginBottom: '20px'
          }}>
            <p><strong>Fichier sélectionné:</strong> {file.name}</p>
            <p><strong>Taille:</strong> {(file.size / 1024).toFixed(2)} KB</p>
            <button 
              onClick={handlePreview}
              style={{
                marginTop: '10px',
                padding: '8px 16px',
                backgroundColor: '#28a745',
                color: 'white',
                border: 'none',
                borderRadius: '5px',
                cursor: 'pointer'
              }}
            >
              📊 Prévisualiser
            </button>
          </div>
        )}

        {/* Settings Section */}
        {file && (
          <div style={{ backgroundColor: 'white', borderRadius: '10px', padding: '20px', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' }}>
            <h2 style={{ marginBottom: '20px', color: '#2c3e50' }}>
              ⚙️ Options de personnalisation
            </h2>
            
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
              {/* Theme Selection */}
              <div>
                <label style={{ display: 'block', marginBottom: '10px', fontWeight: 'bold' }}>
                  Thème de couleur
                </label>
                <select
                  value={settings.theme}
                  onChange={(e) => setSettings({...settings, theme: e.target.value})}
                  style={{
                    width: '100%',
                    padding: '10px',
                    borderRadius: '5px',
                    border: '1px solid #ddd'
                  }}
                >
                  {Object.entries(THEMES).map(([key, theme]) => (
                    <option key={key} value={key}>
                      {theme.name} - {theme.color}
                    </option>
                  ))}
                </select>
              </div>

              {/* Font Size */}
              <div>
                <label style={{ display: 'block', marginBottom: '10px', fontWeight: 'bold' }}>
                  Taille de police (en-têtes)
                </label>
                <input
                  type="number"
                  value={settings.header_font_size}
                  onChange={(e) => setSettings({...settings, header_font_size: parseInt(e.target.value)})}
                  min="10"
                  max="30"
                  style={{
                    width: '100%',
                    padding: '10px',
                    borderRadius: '5px',
                    border: '1px solid #ddd'
                  }}
                />
              </div>

              {/* Font Weight */}
              <div>
                <label style={{ display: 'block', marginBottom: '10px', fontWeight: 'bold' }}>
                  Poids de police
                </label>
                <select
                  value={settings.header_font_weight}
                  onChange={(e) => setSettings({...settings, header_font_weight: e.target.value})}
                  style={{
                    width: '100%',
                    padding: '10px',
                    borderRadius: '5px',
                    border: '1px solid #ddd'
                  }}
                >
                  <option value="bold">Gras</option>
                  <option value="normal">Normal</option>
                </select>
              </div>

              {/* Alignment */}
              <div>
                <label style={{ display: 'block', marginBottom: '10px', fontWeight: 'bold' }}>
                  Alignement
                </label>
                <select
                  value={settings.header_alignment}
                  onChange={(e) => setSettings({...settings, header_alignment: e.target.value})}
                  style={{
                    width: '100%',
                    padding: '10px',
                    borderRadius: '5px',
                    border: '1px solid #ddd'
                  }}
                >
                  <option value="center">Centré</option>
                  <option value="left">Gauche</option>
                  <option value="right">Droite</option>
                </select>
              </div>
            </div>

            {/* Action Buttons */}
            <div style={{ marginTop: '20px', display: 'flex', gap: '10px' }}>
              <button
                onClick={handleBeautify}
                disabled={isProcessing}
                style={{
                  flex: 1,
                  padding: '15px',
                  backgroundColor: isProcessing ? '#6c757d' : '#007bff',
                  color: 'white',
                  border: 'none',
                  borderRadius: '5px',
                  fontSize: '16px',
                  fontWeight: 'bold',
                  cursor: isProcessing ? 'not-allowed' : 'pointer'
                }}
              >
                {isProcessing ? '⏳ Traitement...' : '✨ Embellir le fichier'}
              </button>
              
              {downloadUrl && (
                <button
                  onClick={handleDownload}
                  style={{
                    flex: 1,
                    padding: '15px',
                    backgroundColor: '#28a745',
                    color: 'white',
                    border: 'none',
                    borderRadius: '5px',
                    fontSize: '16px',
                    fontWeight: 'bold',
                    cursor: 'pointer'
                  }}
                >
                  📥 Télécharger
                </button>
              )}
            </div>
          </div>
        )}

        {/* Success Message */}
        {downloadUrl && (
          <div style={{
            backgroundColor: '#d4edda',
            border: '1px solid #c3e6cb',
            borderRadius: '5px',
            padding: '15px',
            textAlign: 'center',
            marginTop: '20px'
          }}>
            <p style={{ color: '#155724', fontSize: '18px', fontWeight: 'bold' }}>
              ✅ Fichier embellissement prêt !
            </p>
            <p style={{ color: '#155724' }}>
              Cliquez sur "Télécharger" pour sauvegarder votre fichier amélioré
            </p>
          </div>
        )}
      </div>
    </div>
  )
}

export default App
