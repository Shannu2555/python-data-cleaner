"""Download the UCI Online Retail dataset (CC BY 4.0) into data/."""
import io, urllib.request, zipfile
from pathlib import Path
URL = "https://archive.ics.uci.edu/static/public/352/online+retail.zip"
out = Path(__file__).resolve().parent.parent / "data"
out.mkdir(exist_ok=True)
print("Downloading", URL)
zipfile.ZipFile(io.BytesIO(urllib.request.urlopen(URL).read())).extractall(out)
print("Saved", out / "Online Retail.xlsx")
