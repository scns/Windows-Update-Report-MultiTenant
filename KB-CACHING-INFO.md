# KB Mapping Cache Configuratie Voorbeeld

Dit script implementeert nu een intelligent caching systeem voor de online KB mapping database.

## 🚀 Voordelen van caching

1. **Performance**: KB mapping wordt eenmalig geladen per sessie
2. **Webserver ontlasting**: Minder requests naar de online database  
3. **Reliability**: Bij offline/timeout situaties wordt gecachte data gebruikt
4. **Configureerbaar**: Cache duur instelbaar via config.json

## ⚙️ Configuratie opties

```json
"kbMapping": {
    "onlineTimeout": 10,           // Timeout voor online requests (seconden)
    "cacheValidMinutes": 30,       // Cache geldigheid (minuten)
    "fallbackToLocalMapping": true // Gebruik lokale mapping als backup
}
```

## 📊 Cache gedrag

- **Eerste run**: Download online KB mapping → cache opslaan
- **Vervolgaanroepen**: Gebruik gecachte data (sneller)
- **Na cache vervaltijd**: Nieuwe download proberen
- **Bij netwerk problemen**: Gebruik verouderde cache als fallback

## 🔍 KB Method kolom toont de bron

- **"Online"**: Vers van online database gedownload
- **"Cache"**: Uit locale cache geladen  
- **"ExpiredCache"**: Verouderde cache gebruikt (fallback)
- **"Local"**: Lokale KB mapping gebruikt
- **"Estimated"**: Geschatte KB informatie

## 💡 Best practices

- Cache duur: 30-60 minuten voor optimale balans
- Timeout: 10-15 seconden voor responsive ervaring
- Monitor KB Method kolom om cache effectiviteit te zien
