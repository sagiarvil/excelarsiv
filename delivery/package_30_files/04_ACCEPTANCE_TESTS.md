# Terminal ve cURL Kabul Testleri
1. H1 Tekillik Testi:
curl -sL "https://excelarsiv.com/" | grep -E -o "<h1[^>]*>.*?</h1>" | wc -l

2. Kanonik Etiket Doğrulaması:
curl -sL "https://excelarsiv.com/" | grep -E -i 'rel=["']canonical["']'

3. IndexNow Testi:
curl -sI "https://excelarsiv.com/7c6b5a4f3e2d1c0b9a8f7e6d5c4b3a2f.txt"

4. Agent Card Doğrulaması:
curl -sL "https://excelarsiv.com/.well-known/agent-card.json" | grep "AgentCard"
