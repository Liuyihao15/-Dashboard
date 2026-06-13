# Write the full V3 html with __DATA__ placeholder, then run embed script
cd /Users/edy/Desktop/霸王茶姬看板

# Backup current index.html
cp index.html index_v2_backup.html

# Run the embed script
python3 _embed_data.py

# Check status
echo "---"
ls -la index.html
echo "---"
head -5 index.html
echo "..."
tail -5 index.html
