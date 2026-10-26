# ==============================================================================
# APLIKASI ENTERPRISE GENERATOR RAB & LPJ DWP KPPBC TMC MALANG (V5.0 PRODUCTION)
# Arsitektur: "Fily" Modern Sidebar Admin UI + Zero-Loss Template Cloning
#             + Smart 3-Mode Financial Engine + 1-Click RAB-to-LPJ + Google Sheets
# Aturan Koding Mutlak: 0 Backslash Character di Seluruh Baris Kode Python
# ==============================================================================

import os
import io
import json
import base64
import shutil
import tempfile
import subprocess
from datetime import date, datetime

import pandas as pd
import streamlit as st
from PIL import Image

import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except Exception:
    HAS_GSPREAD = False


# ==============================================================================
# BAGIAN 1: KONSTANTA, ASET BAWAAN & KONFIGURASI HALAMAN
# ==============================================================================

st.set_page_config(
    page_title="DWP BC Malang — Enterprise RAB & LPJ Suite V5.0",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

NL = chr(10)
TAB = chr(9)
BS = chr(92)
BULLET_CHAR = chr(8226)

DEFAULT_SPREADSHEET_ID = "1_I2f5Gb_F2hMCIIpwwjLF_LpWeCSlWc_E60rYh9hTxc"
SHEET_DB_RAB = "DATABASE_RAB"
SHEET_DB_LPJ = "DATABASE_LPJ"
SHEET_REKAP_DANA = "REKAP_ITEM_DANA"

COLS_DB_RAB = [
    "ID_RAB",
    "WAKTU_SIMPAN",
    "NAMA_KEGIATAN",
    "LOKASI_ACARA",
    "SUMBER_DANA",
    "KET_TABEL_DANA",
    "TOTAL_RAB",
    "CATATAN_NOTE",
    "STATUS_KONVERSI",
    "PAYLOAD_JSON",
]

COLS_DB_LPJ = [
    "ID_LPJ",
    "ID_RAB_REF",
    "WAKTU_SIMPAN",
    "TGL_LAPORAN",
    "NAMA_KEGIATAN",
    "TEMA_KEGIATAN",
    "KONTEKS_HUT",
    "TGL_ACARA",
    "WAKTU_ACARA",
    "TEMPAT_ACARA",
    "SUMBER_DANA",
    "KET_TABEL_DANA",
    "TOTAL_DANA",
    "PENANDATANGAN",
    "PAYLOAD_JSON",
]

COLS_REKAP_DANA = [
    "ID_DOKUMEN",
    "JENIS_DOK",
    "TGL_ACARA",
    "NAMA_KEGIATAN",
    "NO_ITEM",
    "KEBUTUHAN_KEGIATAN",
    "JUMLAH",
    "HARGA_SATUAN",
    "HARGA_TOTAL",
    "KETERANGAN",
]

TEMPLATE_LPJ_FILENAME = "LPJ_DWP GABUNGAN DI BLITAR_2026-1.docx"
TEMPLATE_RAB_FILENAME = "RAB DWP GABUNGAN DI JEMBER.docx"

EMBEDDED_LOGO1_B64 = "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"
EMBEDDED_LOGO2_B64 = "iVBORw0KGgoAAAANSUhEUgAAAFcAAABECAYAAAD0iqlZAAAABmJLR0QA/wD/AP+gvaeTAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAgAElEQVR4nNW8d5xTVf7//zq3pLdJJskkmd6YwtB7laKAgmDHrqjoomtfV3cVEQu2xYK6il3Xiqigwke6MnSYYWCA6X0mmUnvyc3NPb8/9ut+XEXFXfHz+L0ej/tHknPe7/d55txzT72EUor/v+iZJ87z6bI+X3jd1XTT/3UspyLm/zqA71RPiKzCvyb3p35/kBDF2XPONtTszXzvJ9NcQxSPlhJHlBJyeqL8dfrd4S47gyi+/3noK4RfdZPplv3P3ue+9/OOpp/KN/V8WAdcbUBaVP5UmsziRXc77lze9s5TN9V88zQZ/f3fBu95hQ/d/eDvWt7fzdnH980xv/DUZa/Jzrh0YG4OcSRfIPwLD8hve0hxaa9l9LnPV40o1AbDPd0ny/vgxURTWzltB89Gce7cOapPnyNnnizd1TPXf1QxOI+vHF42tEW5aO/rz835Vl5NCgHgjJfv+/u7ttaG15/MWZAUye9S7t/Fyeer7JOdWmuD3Zq7aERllv+dFvS9L4z+uHT41c/ojOnM0ko9EqkBuPo/WnGy/G4vLhk7rjIPkhdFuZkIY+qz95EfA1KO8bRs/Pr9vnCkB0XFBlJUWjDp7eazDxubyJhMj/+FweXZxdachZ/ce33xlyslIj/d5T6tcJOEkGU32Bf7cf62UcPLMhR8As31OzyHPp6yXqfPnq9RCWjtqEUw3Au3vyvUFu9ffzI7GdmKoZFwKyKhdrQ0H8Cg4kE5WwHND9PxDKU6dfUTHC9Qr68Dfn8niossmofWFO1Y8SgyxGiHL9umZUeOmjjbtPKSA7PfIObTWX5yOnsL//PCiHOVeZd/kkqFuFh4ANmODHR2Hk87HJmsq68TuUVWZFh5bN9Vh/37a5e89Ap9+WR2tr1T/HhSpr+nIF8DCTIQIS/c+cfX7bPraPSHaaUbCP/msEl1Fqu5TKVIIxr0IstqhafPHe13BURDRqk+05gPntVhy7bdjd9KpcM2vfB88nSU/7TV3MlXfW7Y3WL7QEyGuYC3B0p5Godrd4ElCdZkkEOvZdHV0YyGhgapKPfI0ycDmwIhymXEHk5qhvY5G7Fn77fweJrx9aa1fPmZuWNTJM39qECv0tTcHdVjDBqhgUhpKOQsokEfsrNz1Dyv0Ks5DgHPAKLBAKZPO3OQuWvHc6eLwWmpuQ8QwtmWzv1kxLip5xIphkTCh2jUDTHlh0wmQavhoFYrwCt4hFPR8IzkF8XRs6n7h3Y+e/WGubbcyg9crh1qg7ERSqUfEmWhV41CX2uVGHDaPr3wzzcvPFkM298jN7O6C1YhHUNKSMJuyUM4mIJOlYX2djeMhmxQmQEtrmDqmosUk0XTX/f/1hzIHQtvXJtbkjvVkalvPbpv66HznZ+9WDUYLfUu0GMPAhcNhnQhQNcAuBhAx2IwhVkgq9R9/P07F74y74J5s+WyRIOcoZEDNUfLKypKGJYlKp1Wk5FMBsFzIpJJPxIJHwhikMsoeBnAMiJEyoHwGeh1xluu/MO60h8Gd24VGXLZH+/eYTS6dSp1PSOkWtDrDGJQ0RnobChMGISqv5x55+3P/jCffA8p/qD54sZcu47Ewh7EohGE/TEwRAkhxoHjdFDJMqA22sHqjAgKSWox67qO1tWxJqMmyDOxXjBKm8vpdgzyvDVj0cU4cWI58GAlpAcfAgUASoHXAOamSpD7jwDxYeDzp2BqPzvuRk1mcYE/LGWRcx6E6hLbtK/KysZNpWIEkUBXioi+YDTcK8ZDPmjUgg9MIpISKSinZiKCPlMiaoVSbeDSaWumVj0IHKMneo2WtrW1EIYFjYRCcGRbiZwH9HolkkKQimIEHJ+GSsECJIVYNEiCkRSNRNTQ6MrTsy95iD/Zvz/yRqJfosW5TH7um7PnVDKtHU3wuGIdG8Y7z1hdSLtOlueZC8moBQ+s2n/syB4UF9kISyXqHvDCYc+HxxUmRpMDvoEw5dUGEhJFGo5FEYtHyEB/H+wOM1UpeUBKobOjBUqeiasUNJgWBSrnJR9PhKgkpcESFr6QlOULhWVVQ0YQT0CQiVAbho+ZQo41dkpp0rmE+3IZjV3qJvNbnm5rnzJ+VEZpiZ2PBNKZMXUcpgo7CIQsMZ1CWqKQiBxpwiORAhJJCiElIujtRaezFUa9AXKep3q9HkFfD0JBUJMpA9F4CslEFCkxDoWSBcszoABC8RQNBATwjA0kkf76p26tQ6/Q4GwZObbolSLi9rJgGAtYrr39p8ACwIZqnBjS0puOhNNsw4keqlLyoOk0vL4IIokkTbkH4PWGkfZ5aUJMg5HxMOhUVJdnRTQWQLezD1qNHIW5BijlrFKp4JVqpRw8x9jSKQHJZBJpIQWrnYVMoYVEFJCQIKzcQFubmnDi6M5373to22oOAD4w09A1t5Jhe54u3C8ETdZUnJBoUITdAljMZhBGAggDwsjAMDLIKA9JFCEKQDIWgpIXkWm2EIVcBZ1ORQMBNQwGLTFkaBEOB5FMixBESqlACSdyYFgGlFGD4RVQKczU7j3wsw8VpQpk+vRrUVu3HjpTPnKL8n70IPu+Nrto9M0nztmQW1R1roxLg2FS0KpV1O8PEoVGjVgsieyibDCEw/FjzZTnCWHSKWh1KjBSGESnoJIUI0QEYskEaIoBK8ohyTgwkAAxBSqmkBYZtHR1QUjJqUpvha8vgKYT+z/664stiwDgX0G+5aDdu917Jte0/K2mMNeiycyUAzSKluZesCwgSRKSQhoSIwOnUINhFBBBQKQ0srPNJEPH0fr6WqTSBOk0hdGko4lEBIQjYHkZOFCA46jEyMCwDHgVgSrNIuqND+hv31SN234aViKYFf9me5O0a+8xNqdQj3Ak5Rg66ufwAub+DY+S4pFz3d5+IuMoFEo1AuEQzeB4eIJuwqlkVBQkGPRaJOMJerTmMChJQKUCBlcUIRZL0QFXJ5QygrA3AldaAEPS4BkWDEtA0xLSaRm0GcUIiAQeVz+NBPu+2Lyq5crl5J+9hH+rARP+Mb5l9wx+atDq2ZNrUcoYSmHXGaFUcSCEIhKLIy5QMDIFWE4JUWKgUWjh8wVoR2s/+vrakF9QAq0uEwxD0d7RCqVai5QoQaQSGI5FNCGAYRgIaQHpOA+LUtOQS2n850BNuuyZMY3H9zCzZl0LlT6FffuPWtY+/UjeBXff3/lTecaH0fdtKEYlyhCW5xGOxsEp5Ygkw5BrZDQY90GMUWh5EzKzHFBwDMRUGO3tx+Dpd8Gg5zFmxHBQMYx4LAghGQZD0lDKePAch3RKRDzBgZGbEItHxDxpx6NfrQg9vOcxKn0Xw7/BvbCfaM6YueAf9pxCLpnoQX9vJyAFYc5UQatRQkiJiCckSDEZxDSPpJAmx90+qtMZSNXgEbSkuAhpicXhow1IJARiNVtoMiUQjuMoJQDhWBCGA6UUVGLBcDxSjKJq00VEf9YaGvwhoPQFArvAufxPi2+Z8/CR+uNEqclBv7sNjvxhmm/3OWtfnrf0vM1fLP/mZHD9cyaMUCTkjFKhAENSaGtrRBopeDwDKBlUSsREmqplaoJ0jLY2dSEaCWLCxFGwGOXw+Xpw/MhhuJ1qGI0KkhJCNC0moFYqIKk1YBiWREJR6g2kITESzFmlSVewcmzsjD33AXj0uxj+rZ+bup9wS4Ozu+edM8Wq5H2oq/kaVitBJNSNkN8Ps1kPDhqwRAebtRiBYAR6gwpKjRbpFAeXK4ABdxRpcOBkChCWQSKVwIhRw9He2YG0SOELhGGx2CGmCCyWHHi8CdTt2/HRnS/svkz1vWCuGwwrN+qGD86df/EZPT1d0Bpl2LN3G+w5FugzDdBoMhEMxNOH6o682G5y3P3N8ofE7/I+P4Rkm66/qdZq1psyjVo0N9ZDJueQTMaRlWWBTKmAQqnGwQP7YDHoEA0GIYki1Go5jHoVVCoGajWFWgV4vV2IxLzo7+uFIIgw6M0gkEFIsVCqbYjRDBQNGoWuHh9Sri9uWXBfx0snhQsAFcFlY25b3bN75HA7w3MDcDoPwWJioJJLyNDp4XMn4HcnkGUshiCIYGRp1B8/RmIRSouLh0BvdKDXOQC3z084uYxaskxobGlCcXEhQpEocnOK0N7Wg1AgiaLiSlCZBuG4CDYd+PSGS/92ZQ+l8cl5ZMaca+5dp1DqVVl2B8ZMGI/33/8HGDkDhZLD5DMmoqO7DXWHj2HE8Ak4fOjw0Ucff3g2jVLn5ytlVb3Jq7aNHz3M1NfViGxbJhgiIsuciZ7eLsSSCXg8HqLT62kw4EV/XztGjhgCvUqHI0cPw+3sRk6OGXabHqmUH9m5JsKzEhVSCaQpC57VIRajiEUowGVAbS5BW5cf+/fva77sgT1lBab/BXrSEdqWV+f8sS9hWDlimJ1tatoBpSyKgK8LGrkSHFFBTHIwqB0I+IMgMgFqjQZKhRGxGODxxpEUJKh0OhiMBnT3dYOTsYhEQtDqTSgsGITuLhcqykdg997DUOjNCEXTGFI1HF2dPR1dXa52W3bZ5IM1R7lF1/0BVpsdz6x6Ho88shw33nITgiEPmlobUFBcAPeAF1dffgOJhKI0EQuGbltyde2Xn68dS5JhhUaRhlpBEPA5oVPLoFXL4XT1QJ+RAbenHzK5HFQSYDIo4ezrBkcZFBTmQ6tUwOXsQMDfC5lcQFqKguUoJEmCkKKApADLGSDnjWDlJhhsFajeddBZtXrbmJk+2vt9juyyZct+BNcw9MoD8q54x8adXbMslgxereQQDAZBwMBksMCgM0KnMYJCAiUiFHI5hKSIYDACjuWRaTaDZwkZ8DgxbGgFwmE/7A4b4vEYjtbVY9z4ydi79yB8/jC6+twwmCwYGPBjy7ZqQ3nFsAJKOea6xX9AT58T/mAIH69Zg1tvvxULL12IceNGo7O7A2fNPgtbtmyDmCSkqnII/D6P3NM/kL99y2Zu5IghqD9agwMH9oJjJCSFGPwBN/rdTlgsGRDTSTAsoNHIkRIisNst0KgU6O5sR3d3G5RKDg6HCXq9CnI5C4VCBqVKA16uBifTwaDPgkKRgUicpTu2bjvwIq2d+cK3tO+HHH92buHQTGJvW2BfbbSWz1ErRMJBJDxhEQ3EIWc1kKgIf8BFQ9EIMeisyC8oA03LaEenE4mUQMzWTCiVPIKxMA7W1NBZs+eTltYe6vcmYHcUkT5XiFaOnEo6nV7a0d4Nk9FG+gf89Oab7yC+QAw2ew5WPPEUguEQHn50Of7xwT9oSWkhzFYj5s47B9POmI4v1m0hJUWl9IXnV5IP3nsDV192Ad5c/TzOP/cs6na1E7ezDQ5HBkwZGurzOokjOwtd3W3geZZazUbicnXQDL0OTJoQn8cDKZ2kBr0KLBFJMNgPa5YBopSCWqenLKsmsSiPRJKjvV3uQAl/ZPnRO6KrLqc0fTJ+Pwv3lRuJauw5pluckTFLkQqpUskADCoViUWigARKRZEoVITqdQZiMGZRMcWS7s4B6vUHYLPZSGlZCWqPHEJjcxMdN2kqurpc0BvshFIl9QcEwitN1OlPIxhLA5TDhRcsJB5viL7y6lu48MLLyM7de1BUXEqbm5sxeEglmTBhHE1RAZYsM+YvmItLLr4M+bnlpK72CI2EfDhzxhRCUmH0u9pow5GDOHv2ZLS31oNIYWIwKChoHEWF2aS3rwuJRIQm4xFSWJhNmxpOIJVIoaJ0EBy2LPh9AxhwdiAlxkAYkQRCQSpXa8HyGYjHZGBYXerRqzW3Pjtq3bZxx2nzT/H7WbiPl5F8cd7447POGqNIRFxIRPqRk2UEQwUIiThEIQaNQQ0hKaLf5cNAvx86bSZycnKRTqfR09cDmYJF8aAy+EIRnDjRDo0+C12dPpgthThc3wFe50Be0RBMmDAJoUgcd991L665djHkMhV27d2Hmto6VFZWQqlW4MMP38f4SROgVMlx8x+XYMbU2bBYHOTyy6+k2XYL/B4XBpwdmH3WJDTW18Bk4GE0cDhxbD8IiUOjBKqqSiDRJDhGRGdXG3RaJcpKB4GKFLX7DyLg96KyYhAqBhWA5SSIUgrO/n7wCh3AGBBNyKEz5KGpaQBfrftiyyOfuWeZuZND/Nlh5L0NtOO9h3VL3O7hbzis2ZAxLJJCHD63E5IYh0rBouVwMxLJFDJ0ZlRUVIBn5Ojr60UwEoZCIYMxwwCffwDd3QOwWMwIRyny8wvR0enBrNnnoMuTQjCaAiUcGpva0NHZg75+Fy66cCG0GUZk5+RBq9fBYslEXl4eBrwD0Ot18PuDeOvdN2HUZ8GoN+LokcNY+syzGFxejFEjhiArOxf9fS3Iyy+BQq2DQqaEkPSiraMVCjkwbGg59IYKnDhejxMN9TBo9RhUVoRk3AG/px/ffPMNGIYiO9cBfyQCY6YavJJBJA4ISKFvIBxVL3Zf9lNgf7HmfqeGddMX7WvV/33E0AJZzb6vkWNTIZ0KIp4IIy+/mHK8gqgUWgR9YXS0dlGWZZGX7yBagxauARdAGPQN+KkENTFmFtLnVn2Eu+6+j5QNnghdVjE2bdtDNRodGTJsJByOfGzeshlDqoZDqVRDoVDgoksuxtq1a6FQK6goCsjKsmLbtm2ort6FhRdeQbq7e2lJcT6ZOmUS9u3ZhRnTJ0Cn5uhdd9xA+rqbMG5sBT7/9C0KKUhkXALpdBhKOYOhQyoABnTA1U+c3b1QsDwspkzkZtuoXMaQo0ePQKvXIBBOwOmPoaRsLOTqfLr2s687dow/Om7PlXTg57id8mR53XIytskw43VzprYiy8wQnglDp1PB6QnRAXeQpJJpyIkCDCFUr9GC40G8AS8EUYDb46XeQALFZSPIqHFz6FcbdqGwcDiROD1du3E32V97jK58+hmseuEFUlIyiN5//wOEZXhs+J+vsWTJEmg0GjQ1NmLcxIk0KcRRe7AGAJCRacSEMZPI8uWP0BHDqkgoHMHOb7fhhuuuop+vfR/l5XnkzVefhdHAIsehoaYMFtU7NxApFUJOdib1elwky26jeq0OapmKCLE4ejo7KBVFlJUUwGg2knAkAqXehI6+AO1zx6KFOZq3Dpy94f6byI9HlP8xXAAYTwl382PlR7V6WVla6KVun4fItRYqSTJiNdup2WAi6aRAU8kkYvEQGXD3wWa3Q2J4KlIZNIZsZGSWoK0zhI6uEAnHWDp++lyMmTANq55/Houuu44ADF2zZg0efvhRIooi7r77btrY3IyNGzfC09+PO/50J9555x1MnDgRM2bMwD133kuSCYHqtDqyY8cOOmx4FdlVvZ02NxyFjE+hr6sBPZ31KC+3YUh5DsyZHGlqOAT3QBc1ZxpIIBimPMshK9NKskxmSqQ0vO5+uJ3dCMfCiCYF4igopr0DMfrn81tm+SrollPl9auXed54eEi7SpnKUymC8Ae9KCgfi3AYkHNyMBQIef1gIUKnU4PlKA7V1sDqyIXBmI1oSgl/iIfFVo5jJ5wwWgpRUjkGTe3dqD10CIRl8eX69ejt6QEATJs+Ha3t7ejr64OYTMKekwOZnENHa/s/g6EAKAOHPQdOpwt5eXkQhCTOmTsHerUMdXX7cNcdi1F7YBuy7WpQ0YuRQwvQ0XYYTY2HkWU1w+7Igc/jgxCLg4gSbBYLdColerpbEBcSkKlUkGtMaO0NQBE5cdk1T9APTxvc3Z8sjvR271XptSHkFOahu5/CORABz/Aw6QxgRBEpIY5kPAyf343svFzEU2kYMvOQkVmEtu4E5Cp7SkrLPudU+Rc8+/fXmLLKoaje+Q0OHTqIP/3pHjgcDuzfvxfvv/ceVr/6KhobG2EymeB2u3G4pgZWqxUlJSWIRuLIMBiw89tdmD5zJqLRKM46czbefusNCKk4eruaMX/uNGcq1Hxw7PhR84K+drh6jsNmVSHLosHhwwcBwqOwsBh5NgdaG5qQikWhUcuRjIdQWloMTziMYCINf4xDY8POd+5dGbvmVFn9bG/hZEoKaag0GoiiDy0trfAnDaBEC5s9F7l2O5LRCMJ+D1IpDXRGPSRK4XcHQWRRKHQUNbVHhCHZ782sv9u/d2Dx9RNHjh5vP3f+AtTVHcE33+5CPB7H8RP1mD9/Pt55511g0XX46qsvodVq8c477+Cuu+6CRqMBz8lx4MABEIAMGTqYFhUVwdnXj7KyClx19XVIJiKgUhKPPPjHtz/f3vTXj1aG3snIdFw+ZOhIpBIDiEVDKC2twMFDh+Fx+2E1ZEGt1sGU5QCoiH27W5BOUyQkiqhIYMgqQZKYf3Iv28n0q5fWQ6EIysoqSFVVFdEbDBCCx/5+21V7L9y4aSvWrvuCrP10Hb7dtRfBUIzY7DmEEo7YHTng5Wo0t7Slbr2Uu/Kiu3zVD1Eqjh0/9ZWRo8aizzWAtZ+tw6atW3HNNddAq9Vi1JjR5PP1n5OvvvoS3d3dqKmpIQaDATu2f4tYNIH33/8Qx48fx/bt26HT6XDRhReSUCiExuZW8vEnn+L4iWZs2rxdWvqHc17MUFLadf/Gq/vaD3wmSQQ+bwgHDhxCWgSZMnk6hCTF/oO1pLbmKOnpdiHg7hu4f5zrJmevK11aOgg5efkkJ68ADrv5RyvUvwncGPnnzkGaTiZ7OrsQDIZRkFeMxv/BrrP6UevzepGbn4lps8dh/BmjoDBkoLHDhzRnA68uRlNLf5c58Nl485ita76z+dS1Vzyx8asv6VtvvIaZ06ejp7OHvPn623jq8b/h0kuuQI4jF/PnzcOyh5aisCgf3T2dePGlVbj22mthtWZh+bJHkE6IGFRYihuvX4z7778Pt9+6BH09HWhtbYZMJmufcfXKXgB4jFLpelJ30dbPn3s4HJMnZ865Asm0Cr0uLxieQ8XgCpSUlcDp7sXb7+6U7ToP74FJxRUqJRFFCX6/H3r58T0AkIic2oa+U0p0c+9+0z2XTNmy/mYyP57wCz6/hxqNmTQwEMN9Y+G99l5UXn3pLCSEVgqVE3HeA0GhpHrHSNrer/PWHu578OnQ7vLZf6U137d7jELgGfKOSadBW2MzPnz3I+px+TF5/HSMGjaeXrHwajq4cgjCAT/On38e3b17J8ACap0eL/79FUycMA0txztw1RVXIx4I0Q2ff4bezib66afvY+LEkdBZdLd83592BZXue0V4cLLu46FbdzZupsp8ITM7H2kmBcKlaFaukbq8HRg9OVOvSSPBcdGAJKWpICapzqDFxh1x07IVxrtzRy5/NdyNX9ymekpw5c+P8DtybEV55z39mccTsAqCgI72Lsh4JW4dhcD696YwPEcxfGQZ+gZa0OHsxLr/+aTB2b1i8rY7V+ff9fD6h3tWnnwpJ9S+8bZLF17gDwZ9SCYTePWVN3D1lYvw2qtvYmDAg8bGRkiShOKyfIweMxK7du3C/PnzkEql0dreBkqBXGsesh12bNu8CQcP7cG582Zj7Wcf+zYGjp50lcJxFW267YFPZ3363JODQm0vzu/p7Y3FBAG79uzGeRdeAIstizw6DWRIhbWT5XikUmls3/4NzNmVf50x9cEnF10/pFmbg1/sCZwS3JVPcNKhrR897g0E4cjJgz07F61tnejq6sHNK6F48plv09F4ApFwAokYhcWcg3tuo3+6/Ca66xOCyM/Z/sfG5tDaT9e+9+QTTwKQUFdXiynTJkFIxTF1yiSsePxR6PRaMFTC4UMH8dDSB6DgZQh4PZg0YSIuvvgi5ORlY/+BA2B4BsNGjCBWWxZUGtljO5a9lfg53y/voB1n3Ui/sJiynCqVDYkEB47TQKI8HvWAaWru4b/+eiuyrNkoLipHjr0Ie3YeSK1XHnnhVLidcptbcACrP1vzkXNw1XCAsoQyLOwOB1pmlj/7cSOqcvIKUH+8lUhUiUQMdOEsHD9V2698+PpSrVaDY0ePkdWrV+OBvzyAO2+/Aw888FdUV3+L5cuXwZFtg92RhQMH92PFikcgSSmMGT0SnT2dEKUUDtcfxmtvvwkJFGqtJmpb9daPduL8lI7Vt3v6+oIoLhmKrzdVQ2/IwsHr8YLCOHzk6LETSEd7L6jIorSoAmLE9dLxW+jPVpjvdMpdsadYKlUz5KxQOHpYo81gCwoGwRsIYsiEaUO3f8MOCQQjUKtNVK5Uoa3NlfzkArhO1TYN0MAtS+7Y73K5xwwbOhpbtm5CPB7B+i/WYOz4oXji8cewZMmNJDs7m5rNWWCIHN6BEJ5+6hnU1dZBpVIgmojgvAvnYtyEcXQg6ln1mEIm/bLnf6qIeJ4Mh8jarMxMuD2HUFLEQtTPuL7YZkMokqRymRpWazZaOtztgQe33nWqdn9VV2zSB/TYzg2vfhiNJVFeOZyIaUp6nF0YPmY4s2XzDpQWDyVmcz5YosG4k+4R/2ll5+bdW1xcjEsvXYg9+3ajq7uDlJWXkgkTxuPuu+/A0gcewZVX3ITZsxbgrDPnYP/eani8bogQseHrjdi8dRP5dudODB05QvxwwqYnf41vjT5LZ7cVoLO7n1RUDIOz34NEKgWP34c+l5OMGjUKYkpMN777+nVPcvSU/7Rf3c/1vRi67pvNH3/T2NQmFZeWU0pEaPUKqtbpEAwlqUqWCYe9VH7+fmQsfZMopjbuyDwVu/fdd8duuVzWpddrwRICv99Da2v206NHavD6a6uRToHmOBRgAJw3fyb6eruQSkbBy1g6cfIETJs+jd795z/h5ddf/KrjjMOBU/I5959b+jcf5/4gk8kQDAapSq2A0ZgBk9kElmcwbNhgWr1rh3Bs3dPXL9tLd/waVqc8/P2IEHYQoM0WETy4Ayy/C9fXBdQPjp04x9rj6kdhSSH5csNXmLfgEqpWZeHlBQ/ka88beUdDG1/+ycY9s08pGKK88+iRI39bv349vJ5+qDUsFHIJCiUDCgF9Pb3Izy+G2+mjSoWOJARKeZWWyBRypJHC2Amj6TX3Xn1e17d9637JV8fBTtuLb9/aOipbLTwAABARSURBVDu9ftwuMnHNOXPnlLY2H0J762FMmz4aep2ctne0onrT7h2pl3DHExm07l5CmHX9UJ2wnFqbe8pw/7p45iKN2b5KxXkiad/Gj+604NUHN6DXVIQFglV+jy13eMmAJ0SMRjsGDx1NLbJ9t62vzn48lpCJtz7wmlHO/vLtNPuPG+RVsl0bcnKzs0NBD+QyEdHoAMswtCAY9pC0kIJOZwSReKSSSKjV+nqjPVvf6xqARES09DRuLc1ce8sjT7C/6Oua88nC65e88P5rK+/fe+vNf8jsdnYXb/jqI1x5+QKIKZ+wd9vWbe+Pwz3XFSFAV2P6uIsuumb7/uCIATcrXfz6Rsf4X9gl9KvgnnXLRvkk1bsNM6cPzxdivbSn+ygtztcMHK5e/7HwMJ6d+RpGbe0f9sHI0ZMYhpVh3ZcbkpOmzpYDWgyiD1cVX0qPnZKjH+gvImE814/9JLegbIFOo4VMpoSr1w0T8Vz6xxVffPSf2ASAG68ZvPOmGxZPPLxnO+1uO44zz5xCxFQYG778yHNbEtP1V2LkR7XMLdnF06s0+hxZY4sXWY4q+sm7z97+6ueRVafi45Tb3E0vzEmONXwwp6WlUzJkWGGx5hBWlmHNr5ryR/Kw+ag5D9npREp0dfcBooDxYyrkUsoJq1nCqg2mpf8phMc4Kg0ZTNYVFuQg06SHTquCzWpM/2nUwb3/qc37lxKuonzQ8IG+JpSVmEkqMUCS8QEYdDxkLEJPiXj9QHTGm0UVs0byMr1MrtCB5RTYf/Cbg6cKFviVD7RZf6ENrQ2HPm/v6EFaIuju6kNBQQHOX3ChuqZ/3t90KoOsrakJvv4eDB7kQDzSiUi4FY5s63gA+JgQ2a8FAQBtDYdTyUQcXR2t6O1uB5VEcu+TLvWvtSMQQpIiYeebUDa4vEDV1XEERPJjzMhi9HScwPGjBzFu3JjCRYsvGT24cjjSKRnCEQnO/hAkIk9//V71KXfDgP+gtzBm1d5rj584EVCpjcgvLEVXRzuaGo5DTIooKylDtsWC9pZjEJIDJNvOg0E/cvP0pvuXEsXro80H6nHGj444/ZIO1SQi8XiIUqShUsqQoVfGr9qP3p/LQ/DjI6rP3Fpx1UOLzU9+tReFRj0HtTwNc4Yc+TlGxCNuqGQcssxZcPa4sXf3EUSjLCoqxhKNxg5vT+O71V20+tfE/avhnkNo6NbB2ycertnXEwpFkJebC61CBYNGCyUvI0adFixNQRQ8UMji0KgTRM7HmPnLwYwbP7Zy3a15K79vL50gpHyM5dmpt+0w/JRPzXbs5FkkM0162CxmMAwjb5sI/U+lv+22i//CsLjm+99d9twcHWcYsdKQkZl3ogZMOOCEI0tPOJKAmAhCybPIy81GwBdEMkGRl1MGsykfdTWt0pGDe9b/kd9//a9l9R8dldJeRU9cdvmuwT1N6z9w9faB51iaZbHB5/FS0DRsViMYxKnP0wZJ9FEGIdw4ChhSXgqNKfNaYev6Id/ZYhWUzpoxb/oY5ZpNYupD9mT+nn4XTH9/PwRBQCQSgauvJ+jbhZPOrSqxa5TNWrzs3Nnjcr7/fXGr8qtBg4oz2lpOQAlATIYgSTEa9LrAQQIHCjkjA6EsrKZs8KwO9XUtCY1r0/Wm5dsu4Jad+uDhO/3H59DMpTR0xX2+y6+Nbqva+fWXn4iJaD0n54IZVhOVq+XwB4JIUx7hqIDMTENsqRyswcST0lI7e9/H7+ySlq75V/MwdoxFqqxUjfzTXY886w3/+Ha+9HWEZCpOMNo1UGUQRFO+xIeA8MN0R54n2TfdsuKrieMruOFDVf+atHlrheXKUWMHTTKb5Kg7ADy6CAMZOgdC7iSQVsCUkQ0hwYNSDdQqW/LY0cbG3Z+99fjAsPXWMx/sePPSn9iu9Ev61cs8P1TrDfRYYDG55Gzra+zqwdB98gbO0g+zvX3uqBmyWDwIY4YZy/76xmahHkUXpbwkEPVg/NR89fO9z718Oy66AgAUpBpGk44Uzq1YsnrpMd3eCWTRuov+t0AuJ2Qkg+P6E+2QUgLavK2Jcunfp/zWXUwGuSfM3jVvdpUpGjpOA93bWgGgsI3oH8oa94LJ4qWtjY24556Fud84PjxoapcSuVnjlF1t+6mi2IJQIN21+/3335K/hA+jV6B1CaWp/5bNb3KCcumCkU+dH532YTiimvfkHWgW0mFnfX0Lvtm+n+7be1isKcIdI+cjL8uiQl62GkRyI0NvuSgPJB8AEgE3OOqDnPOQvIJhV87ek/Ph2R+l/9VE9N+EZCjqEaKiF0odhwkzyo8+Sv73Ns38Gxleb7MeqCi1GHs7j1Ex6iUzRkMEgOvfxrPDh+VrGxp3obDIDLfbY77sIirs/fLlD1QqBUyZFiQECbOmDT9x2QXZPdxL0++sv2j8gUcvHjbsv+Xym8DNsajDRYNGna/M/vOb77bN2pvtKMll0imq4hmkAkcPH3qGOo2t8MV9TrBCGK6248gz6/lLr8afASDpzk5q+UIE+uMYXlWBwWNmXDB986xP3f9vaSl+L9IaFYnLWQ4cI0e4ny36l/MzSclLZHL1vLlnaVKxCIqyC2Ex5OCTlUhdXkfURsPwK5IxgmxbNqxWPSqrTAkA6L6ZPlG99wOppesQ6hr2wCnIZ+0Wpq4uHnPG9aMnjhqcebCu/b/l8pvAfcXy7Qq3s8OfaTCQ8pISVq9Uw2G1kMsvuZg47KNHfvEoWXf7SoxMRf0wa3iSY1bDkaEgDvuoawFAG7XVGlUFCLqjxO/pRUWFhZSOyJ+389biawCALgPhCVXKWTnCnhDMOmE3AJxLCPvK0NKPJo0frNQqKfE6e0l2po3IqUq6thNHT/hQdt7sS1mfK0QIgM72BrJn72ex3r+TufP7squteRIZMbYI9jwT1Bk8seUaiS/YB6OydvmNbb+8o+aX9JvAbXmYCpeZ11zyzY7t6UgoCjnHI+gJoLOpHYXZuaSwfPa8L45PWenv7YaWS2NsZRHkYgjjq0bIAKBgl25N2A1kW3PgGWiH13sMQ4friduUtUK6kfDKx6kU9Ha5xXga6QTQ1x7UA8Cti8deOHnalKHdbYcQ9HYgN8uI9sZmHNtfG5pCafff3rEWR/0pkgiJaG1oQmamBr1uYdgJ4/B1g4bYM0urTCQQd8GenwWBxtHn7qEHard0fNpX/aumLH9Kv9mpdc0f6JZKffuDoVCcFpdUQa3JpHVHGmjDiTZKwIMjLNzOPvj6nTRDo0DY56RSMkKaiUBGbH5xa83efdHigmxqMnLo7DhIqeSko0aXWGavRAUA+N0tNYyogM3kAEeNxQCwN61/R6OgRMbFISVDVK+V0WjAQy29O58AgLa4cWhD/TEa8HipUq5AJBqgixZfSBzZVlJTtx89zg7a0esCq9AgmqC0sak55lvdMuntZfRnl4dOVb/pKwHeW1y3wpDcdPf+muMpXmUio8ZNJ4zMQHiFgSSSBMUl5SQSiZBYPEK0GjV6e1rFEiqjAHBfxpbz3ANNktUsI3I+RsKhLmLMkJGKG+XvAUCOAV9w0ELOaUlFvupgeipRTJkynAsFncSRZSIyTiDx6ACkeLfT/XrncwAwwXTiaHGhjSQTQWI0aLF1y3awjII0t/QQXqaDRpdDpp5xAZwDQN2RE53PX1I/7mnfj7ff/6f6TeE+BSqNWux9ZlHOlpFbd1S3+CJEsuUNhkydBV9ERE5ROfrcfpxoagMFiN954HAXIaSNEDL6PmwPDezYByYJmz0TqVQCkXAYfh8pAABHN7ZKSSU8riD98JPN+U9/Cz63wET0ej06OrqQTqeglBOcp665cwyQ2NRLyO0vYUci4UI46gLDSigrqyR1h1tRXjEVgwfPJip1Ga097BHffnnNa7cv7BrUUETrf0sep+1NIeJSwl25DjOvvnnixUeO+yZOnjA5jyGizO/twvvvbcGMaVbo9BZBrzQf7HH1GSjbwKQIipRy8DIe0OmKYDCPxcdraxN/e+e4auMyUtGmv6WeimEIvuqXKzNad1kG3/luZKAWLHXB7e9CIBiHTmbtZyWVj+cNksFmUH+2eWu+o4BFbn4l0qIOVZUzcexYU6zPua9Fw3VsqHkg/frzlLaeDgan9TUs38kZJWTACvOxR+Y2ObubdZMmViEWHSBWu40GwymwXAoGfRgmIwiLFPV4BHR3y4nbp0dzG5FeuFhbuWfiBx2v3n1VtDA/gwwfFHnu3Xc3njVx/JhyLukldquMGjIlotfrqUJmIz5Pknr9ISg0IPmDtDhcX0u7e+JobxUwNH/EumPvrr1s8X4kDKe58L8L3JwuYr9sxbjqKZPG5Q8dnI1jR6shk6cQCDoBGUVKCECIesDQJHRqEwzafFDkIRTPBPgifLv3eGrvoYPdi264smD48ALS0rA3HQv3MVajmug4DhwTQSjWingiBIaqkEz885xxll0LT6gB4XgCxYUToZRVovkEK25et3vd/Z9uutj6M1vufwv918PfU1FLFlw7x+y95MAx1du8XCxTG8zIyADkWhaCFIHekAONTAGkAEZUQ0rriWuARTAsCtu3r3nv07dffwL6cmfs2kue++SzDVcInnj5JedvrHxhtfTA9Mnnjcg0qaDW5cFgZaDkNfD7gwhH/BAogaOgGMFQGOGYgOrqauo8ln6tqerlh61c+WmvVb9Lzf1OfoGQ2EbM+GC3funQURMnGkxaMhCKwGTMglxSwtnlpf3tbjcbDW3Ojda9N/Sl5J4M+r+d+Tc/e/HBXZu23/Xa3z/RfffdRfKxBfe+lzOzvn/gMrXWNMJmKdASCqRTCegyAF/4BI4c2x9CPPb+PVOwMnkGbfm9yvu71NzvlCGjNGM+tvi0ZEeopa3e7TOUtvf1tw8qCVVL7kTvTLH2q4xlYp0PiJoopXjx3/N3+fqQYv79bS1rkvvaAby64mPyxsT1sCiGYVCLK2O0MauwwBkg13PKAfbGq2MLNNl0x8+80uH0iFL6+1w2qG4onD7inEyc3bZr7tpNH86UPnhphGsrBXeqNqY+OHru5X++YvWppr94Im48XrtUevuNCd4nH8bNX/8DE68qgUP4f3fs6b5+l2bB9gVRPdV/5z6bzVSq5AW+t/cEcuxacQh5fbLqLLrvVO0QgZB4MAKFWX1KQd8jEGbQa9PX2XIqz1FrjZDxGdKh/W2xnFzVE/MvWPHoL1v47/S7vMvxqeBNz1KFvDIspHhRphdF3uZtavv0ml8DFgCojNJTBQsAT8qotMi8bcH+/XVvdXcmY6GQiNHjRmg6uwMPKewk55ct/Hc67TX3vMvzbMPKJrXnF9gEW7b/yUN1b3y99ks0Hdz83886naqSAiFte5H10hqMKCpYeEdBYfn0np5DX958y7pzT6vj09nmhKMgq1bd9I+blkw6dNvDsP1u7fsvXLX7rrj8liVTQ4KwZNLp9HNaCzFprPaKxx5bcN8eAfz/NdAfXoNzULjwkiHVI0ci93T5+P8AJGUIUpjGozEAAAAASUVORK5CYII="
EMBEDDED_SIG_B64 = "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"

DEFAULT_UCAPAN_DONATUR = (
    "Kami mengucapkan terima kasih yang sebesar-besarnya kepada para donatur "
    "yang telah menyisihkan sedikit rezekinya untuk kegiatan ini sehingga dapat "
    "berjalan dengan sangat baik. Semoga bantuan dari Bapak/Ibu dan Keluarga "
    "dihitung Allah SWT sebagai amal kebaikan yang bernilai pahala, bermanfaat "
    "untuk sesama, dan menjadi berkah untuk kita semua. Selain itu, semoga "
    "dengan kegiatan ini semakin memperkuat tali silaturahmi dan kekeluargaan kita."
)

PRESET_RAB_JEMBER = {
    "id_rab": "RAB-2026-JEMBER",
    "baris_1_judul": "RANCANGAN ANGGARAN BIAYA",
    "nama_kegiatan_rab": "PELAKSANAAN PERTEMUAN RUTIN PAGUYUBAN DHARMA WANITA PERSATUAN DJBC DI LINGKUNGAN KANWIL DJBC JATIM II",
    "lokasi_acara_rab": "(DI KPPBC TMP C JEMBER)",
    "sumber_dana_rab": "Kas Dharma Wanita KPPBC TMC Malang",
    "ket_tabel_rab": "",
    "merge_ket_rab": False,
    "catatan_note_rab": NL.join([
        "Iuran DWP Gabungan untuk 5 orang Rp 300.000 + (2 orang x Rp 75.000)",
        "E-Toll (PP) dari Tol Singosari – Tol Probolinggo Timur dan sebaliknya Rp 112.000 x 2",
    ]),
    "status_konversi": "BELUM DIKONVERSI",
    "items_rab": [
        {"NO": 1, "KEBUTUHAN KEGIATAN": "Iuran DWP Gabungan di Jember", "JUMLAH": 7, "HARGA SATUAN": 0, "HARGA TOTAL": 450000, "KETERANGAN": ""},
        {"NO": 2, "KEBUTUHAN KEGIATAN": "Doorprize", "JUMLAH": 7, "HARGA SATUAN": 50000, "HARGA TOTAL": 350000, "KETERANGAN": ""},
        {"NO": 3, "KEBUTUHAN KEGIATAN": "Mobil (Bensin)", "JUMLAH": 2, "HARGA SATUAN": 0, "HARGA TOTAL": 0, "KETERANGAN": ""},
        {"NO": 4, "KEBUTUHAN KEGIATAN": "Sopir", "JUMLAH": 2, "HARGA SATUAN": 0, "HARGA TOTAL": 0, "KETERANGAN": ""},
        {"NO": 5, "KEBUTUHAN KEGIATAN": "Snack", "JUMLAH": 10, "HARGA SATUAN": 20000, "HARGA TOTAL": 200000, "KETERANGAN": ""},
        {"NO": 6, "KEBUTUHAN KEGIATAN": "Air Mineral 700ml (isi 12 botol)", "JUMLAH": 2, "HARGA SATUAN": 23000, "HARGA TOTAL": 46000, "KETERANGAN": ""},
        {"NO": 7, "KEBUTUHAN KEGIATAN": "Sarapan pagi", "JUMLAH": 10, "HARGA SATUAN": 30000, "HARGA TOTAL": 300000, "KETERANGAN": ""},
        {"NO": 8, "KEBUTUHAN KEGIATAN": "Makan Malam", "JUMLAH": 10, "HARGA SATUAN": 30000, "HARGA TOTAL": 300000, "KETERANGAN": ""},
        {"NO": 9, "KEBUTUHAN KEGIATAN": "E-Toll (Pulang-Pergi)", "JUMLAH": 2, "HARGA SATUAN": 224000, "HARGA TOTAL": 448000, "KETERANGAN": ""},
    ],
}

PRESET_LPJ_BLITAR_2026 = {
    "id_lpj": "LPJ-20260813-BLITAR",
    "id_rab_ref": "RAB-2026-JEMBER",
    "tgl_laporan": date(2026, 8, 13),
    "nama_kegiatan": "Pertemuan Rutin Paguyuban Dharma Wanita Persatuan DJBC di lingkungan Kanwil DJBC Jawa Timur II",
    "tema_kegiatan": "Kebersamaan dalam Keceriaan, Bersinergi dalam Kebahagiaan",
    "konteks_hut": "menyambut HUT ke-81 Kemerdekaan Republik Indonesia",
    "tgl_acara": date(2026, 8, 10),
    "tgl_acara_manual": "",
    "waktu_acara": "08.30 - selesai",
    "tempat_acara": "KPPBC Tipe Madya Pabean C Blitar",
    "rangkaian_kegiatan": NL.join([
        "Pembukaan acara dengan sambutan dari Pembina DWP BC Blitar Bapak Nurtjahjo, Ketua DWP BC Blitar Ibu Hasrifah Nurtjahjo, dan Ibu Ketua DWP Kanwil DJBC Jatim II Ibu Ana Lukman.",
        "Kegiatan lomba menghias tumpeng dan lomba sambung lirik lagu Nasional untuk menyambut HUT ke-81 Kemerdekaan Republik Indonesia.",
        "Kunjungan ke Museum Bung Karno.",
    ]),
    "sumber_dana": "Kas Dharma Wanita KPPBC TMC Malang dan donatur",
    "ket_tabel_dana": "Kas DWP BC Malang & Donatur Anggota",
    "merge_ket_vertikal": True,
    "tampilkan_ucapan_donatur": True,
    "teks_ucapan_donatur": DEFAULT_UCAPAN_DONATUR,
    "catatan_ig": "*Note : dokumentasi kegiatan dapat dilihat di Instagram @dwpbcmalang",
    "jabatan_baris_1": "Ketua Dharma Wanita Persatuan",
    "jabatan_baris_2": "KPPBC Tipe Madya Cukai Malang",
    "nama_penandatangan": "Ny. Rini Johan Pandores",
    "tampilkan_ttd": True,
    "items_dana": [
        {"NO": 1, "KEBUTUHAN KEGIATAN": "Iuran DWP Gabungan di Blitar", "JUMLAH": 10, "HARGA SATUAN": 0, "HARGA TOTAL": 675000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 2, "KEBUTUHAN KEGIATAN": "Doorprize", "JUMLAH": 10, "HARGA SATUAN": 0, "HARGA TOTAL": 0, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 3, "KEBUTUHAN KEGIATAN": "Mobil", "JUMLAH": 2, "HARGA SATUAN": 0, "HARGA TOTAL": 0, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 4, "KEBUTUHAN KEGIATAN": "Hadiah Peserta Lomba Tumpeng", "JUMLAH": 3, "HARGA SATUAN": 250000, "HARGA TOTAL": 750000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 5, "KEBUTUHAN KEGIATAN": "Kerudung Merah", "JUMLAH": 10, "HARGA SATUAN": 50000, "HARGA TOTAL": 500000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 6, "KEBUTUHAN KEGIATAN": "Baju Atasan Merah (Blouse)", "JUMLAH": 9, "HARGA SATUAN": 90000, "HARGA TOTAL": 810000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 7, "KEBUTUHAN KEGIATAN": "Sarapan Pagi (Roti)", "JUMLAH": 20, "HARGA SATUAN": 13000, "HARGA TOTAL": 260000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 8, "KEBUTUHAN KEGIATAN": "Snack (Keripik Usus)", "JUMLAH": 12, "HARGA SATUAN": 15000, "HARGA TOTAL": 180000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 9, "KEBUTUHAN KEGIATAN": "Makan Malam (Ayam Pemuda)", "JUMLAH": 13, "HARGA SATUAN": 100000, "HARGA TOTAL": 1300000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
        {"NO": 10, "KEBUTUHAN KEGIATAN": "Bahan Latihan Tumpeng, Support Makan Siang & Kopi Untuk Latihan", "JUMLAH": 3, "HARGA SATUAN": 200000, "HARGA TOTAL": 600000, "KETERANGAN": "Kas DWP BC Malang & Donatur Anggota"},
    ],
}

MENU_DASHBOARD = "🏠 Dashboard Utama"
MENU_RAB = "📋 1. Susun RAB"
MENU_LPJ = "📝 2. Buat LPJ"
MENU_PREVIEW = "🖨️ 3. Preview A4 & Download Center"
MENU_CLOUD = "🗄️ 4. Database Cloud"
ALL_MENUS = [MENU_DASHBOARD, MENU_RAB, MENU_LPJ, MENU_PREVIEW, MENU_CLOUD]


# ==============================================================================
# BAGIAN 2: HELPER FUNGSI BAHASA, KEUANGAN & DETEKTOR KONSISTENSI
# ==============================================================================

def parse_num(val) -> int:
    """Mengubah nilai string format Indonesia (Rp 250.000, 112rb, -) menjadi integer murni."""
    if val is None or isinstance(val, bool):
        return 0
    if isinstance(val, (int, float)):
        if pd.isna(val):
            return 0
        return int(round(float(val)))

    s = str(val).strip().lower()
    if not s or s in ("-", "rp -", "rp-", "nihil", "gratis", "free", "nan", "none"):
        return 0

    s = s.replace("rp", "").replace("idr", "").replace(TAB, "").strip()

    multiplier = 1
    if "juta" in s or "jt" in s:
        multiplier = 1000000
        s = s.replace("juta", "").replace("jt", "").strip()
    elif "ribu" in s or "rb" in s:
        multiplier = 1000
        s = s.replace("ribu", "").replace("rb", "").strip()

    if multiplier > 1:
        s_clean = s.replace(".", "").replace(",", ".")
        digits_dot = "".join(ch for ch in s_clean if ch.isdigit() or ch == ".")
        if not digits_dot:
            return 0
        try:
            return int(round(float(digits_dot) * multiplier))
        except Exception:
            return 0

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.split(",")[0]
    elif "," in s and len(s.split(",")[-1]) <= 2 and len(s.split(",")[0]) > 3:
        s = s.split(",")[0]

    digits = "".join(ch for ch in s if ch.isdigit())
    return int(digits) if digits else 0


def fmt_ribuan(val) -> str:
    """Mengubah integer menjadi string titik ribuan Indonesia (2094000 -> 2.094.000)."""
    n = parse_num(val)
    if n < 0:
        return "-" + f"{abs(n):,}".replace(",", ".")
    return f"{n:,}".replace(",", ".")


def fmt_rp_cell(val, is_dash_if_zero: bool = True) -> str:
    """Format sel Rupiah tabel RAB/LPJ tanpa karakter Tab liar agar anti-potong di PDF."""
    n = parse_num(val)
    if n <= 0:
        return "Rp -" if is_dash_if_zero else "Rp 0"
    return "Rp " + fmt_ribuan(n)


def _terbilang_dasar(n: int) -> str:
    satuan = [
        "", "satu", "dua", "tiga", "empat", "lima",
        "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"
    ]
    if n < 12:
        return satuan[n]
    if n < 20:
        return _terbilang_dasar(n - 10) + " belas"
    if n < 100:
        hasil = _terbilang_dasar(n // 10) + " puluh"
        if n % 10 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 10)
        return hasil
    if n < 200:
        sisa = _terbilang_dasar(n - 100)
        return ("seratus " + sisa).strip()
    if n < 1000:
        hasil = _terbilang_dasar(n // 100) + " ratus"
        if n % 100 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 100)
        return hasil
    if n < 2000:
        sisa = _terbilang_dasar(n - 1000)
        return ("seribu " + sisa).strip()
    if n < 1000000:
        hasil = _terbilang_dasar(n // 1000) + " ribu"
        if n % 1000 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 1000)
        return hasil
    if n < 1000000000:
        hasil = _terbilang_dasar(n // 1000000) + " juta"
        if n % 1000000 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 1000000)
        return hasil
    if n < 1000000000000:
        hasil = _terbilang_dasar(n // 1000000000) + " miliar"
        if n % 1000000000 > 0:
            hasil = hasil + " " + _terbilang_dasar(n % 1000000000)
        return hasil
    return str(n)


def terbilang(n) -> str:
    """Mengubah nominal angka menjadi teks Terbilang Rupiah Title Case."""
    angka = parse_num(n)
    if angka <= 0:
        return "Nol Rupiah"
    teks = " ".join(_terbilang_dasar(angka).split())
    return teks.title() + " Rupiah"


def format_hari_tanggal_indo(dt) -> str:
    """Contoh: date(2026, 8, 13) -> Kamis, 13 Agustus 2026."""
    if isinstance(dt, str):
        s = dt.strip()
        if "," in s and any(b in s for b in ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]):
            return s
        try:
            dt = datetime.strptime(s[:10], "%Y-%m-%d").date()
        except Exception:
            return s
    if isinstance(dt, datetime):
        dt = dt.date()
    if not isinstance(dt, date):
        dt = date.today()

    nama_hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    nama_bulan = [
        "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    return f"{nama_hari[dt.weekday()]}, {dt.day} {nama_bulan[dt.month]} {dt.year}"


def normalize_items_df(records, default_ket: str = "") -> pd.DataFrame:
    """Menerapkan Logika 3-Mode secara reaktif tanpa memicu cursor-jump."""
    rows = []
    if isinstance(records, pd.DataFrame):
        raw_list = records.to_dict(orient="records")
    elif isinstance(records, list):
        raw_list = records
    else:
        raw_list = []

    idx_counter = 1
    for item in raw_list:
        kebutuhan = str(item.get("KEBUTUHAN KEGIATAN", item.get("kebutuhan", "")) or "").strip()
        if not kebutuhan and parse_num(item.get("HARGA TOTAL", 0)) == 0 and parse_num(item.get("HARGA SATUAN", 0)) == 0:
            continue

        jml = max(0, parse_num(item.get("JUMLAH", item.get("jumlah", 1))))
        h_sat = max(0, parse_num(item.get("HARGA SATUAN", item.get("harga_satuan", 0))))
        h_tot_in = max(0, parse_num(item.get("HARGA TOTAL", item.get("harga_total", 0))))
        ket_raw = item.get("KETERANGAN", item.get("keterangan", default_ket))
        ket = str(ket_raw if ket_raw is not None else default_ket).strip()

        # Logika 3-Mode Smart Financial Engine:
        # Mode 1 (Normal): Harga Satuan > 0 -> Harga Total = Jumlah * Harga Satuan
        # Mode 2 (Lump Sum / Borongan): Harga Satuan == 0 dan Harga Total > 0 -> Harga Total tetap
        # Mode 3 (In-Kind / Fasilitas): Harga Satuan == 0 dan Harga Total == 0 -> Rp -
        if h_sat > 0:
            qty_calc = jml if jml > 0 else 1
            h_tot = qty_calc * h_sat
        elif h_tot_in > 0:
            h_tot = h_tot_in
        else:
            h_tot = 0

        rows.append({
            "NO": idx_counter,
            "KEBUTUHAN KEGIATAN": kebutuhan,
            "JUMLAH": jml,
            "HARGA SATUAN": h_sat,
            "HARGA TOTAL": h_tot,
            "KETERANGAN": ket,
        })
        idx_counter += 1

    if not rows:
        rows.append({
            "NO": 1,
            "KEBUTUHAN KEGIATAN": "Kebutuhan Kegiatan",
            "JUMLAH": 1,
            "HARGA SATUAN": 0,
            "HARGA TOTAL": 0,
            "KETERANGAN": default_ket,
        })

    return pd.DataFrame(rows)


def detect_location_mismatch(lokasi_judul: str, df_items: pd.DataFrame) -> list:
    """
    Mendeteksi jika di Judul RAB tertulis Jember namun ada baris item yang masih tertulis Blitar
    (seperti typo di file asli RAB DWP GABUNGAN DI JEMBER.docx baris 1: 'Iuran DWP Gabungan di Blitar').
    """
    kota_list = ["jember", "blitar", "malang", "surabaya", "kediri", "madiun", "probolinggo", "pasuruan", "banyuwangi", "tulungagung", "bojonegoro", "gresik", "sidoarjo"]
    lok_lower = (lokasi_judul or "").lower()
    kota_judul = [k for k in kota_list if k in lok_lower]
    if not kota_judul or df_items is None or df_items.empty:
        return []

    target_kota = kota_judul[0].title()
    warnings = []
    for _, r in df_items.iterrows():
        keb = str(r.get("KEBUTUHAN KEGIATAN", ""))
        keb_low = keb.lower()
        for k in kota_list:
            if k in keb_low and k.title() != target_kota:
                warnings.append({
                    "no": int(r.get("NO", 0)),
                    "item": keb,
                    "wrong_city": k.title(),
                    "target_city": target_kota,
                })
    return warnings


def smart_paste_parser(raw_text: str, default_ket: str = "") -> list:
    """Mengekstrak daftar belanja dari teks WhatsApp/Catatan ke format 3-Mode secara cerdas."""
    parsed = []
    if not raw_text or not raw_text.strip():
        return parsed

    lines = [ln.strip() for ln in raw_text.split(NL) if ln.strip()]
    for ln in lines:
        clean_ln = ln
        if "." in clean_ln[:4]:
            left_num, rest = clean_ln.split(".", 1)
            if left_num.strip().isdigit():
                clean_ln = rest.strip()

        if "|" in clean_ln or ";" in clean_ln:
            sep = "|" if "|" in clean_ln else ";"
            parts = [p.strip() for p in clean_ln.split(sep)]
            nama = parts[0] if len(parts) > 0 else "Item Kegiatan"
            jml = parse_num(parts[1]) if len(parts) > 1 else 1
            val_3 = parse_num(parts[2]) if len(parts) > 2 else 0
            val_4 = parse_num(parts[3]) if len(parts) > 3 else 0
            ket = parts[4] if len(parts) > 4 and parts[4] else default_ket

            if len(parts) >= 4 and val_3 == 0 and val_4 > 0:
                h_sat, h_tot = 0, val_4
            elif "total" in (parts[2].lower() if len(parts) > 2 else "") or "borongan" in clean_ln.lower():
                h_sat, h_tot = 0, val_3
            else:
                h_sat = val_3
                h_tot = (jml if jml > 0 else 1) * h_sat if h_sat > 0 else 0

            parsed.append({
                "KEBUTUHAN KEGIATAN": nama,
                "JUMLAH": jml if jml > 0 else 1,
                "HARGA SATUAN": h_sat,
                "HARGA TOTAL": h_tot,
                "KETERANGAN": ket,
            })
            continue

        lower_ln = clean_ln.lower()
        found_mult = False
        for sym in [" x ", " @ ", " * "]:
            if sym in lower_ln:
                idx_sym = lower_ln.rfind(sym)
                left_part = clean_ln[:idx_sym].strip()
                right_part = clean_ln[idx_sym + len(sym):].strip()
                h_sat = parse_num(right_part)
                tokens = left_part.split()
                if tokens and tokens[-1].isdigit():
                    jml = int(tokens[-1])
                    nama = " ".join(tokens[:-1]).strip()
                else:
                    jml = 1
                    nama = left_part
                parsed.append({
                    "KEBUTUHAN KEGIATAN": nama or "Item Kegiatan",
                    "JUMLAH": jml,
                    "HARGA SATUAN": h_sat,
                    "HARGA TOTAL": jml * h_sat,
                    "KETERANGAN": default_ket,
                })
                found_mult = True
                break
        if found_mult:
            continue

        if " total " in lower_ln:
            idx_tot = lower_ln.rfind(" total ")
            left_part = clean_ln[:idx_tot].strip()
            right_part = clean_ln[idx_tot + 7:].strip()
            h_tot = parse_num(right_part)
            tokens = left_part.split()
            if tokens and tokens[-1].isdigit():
                jml = int(tokens[-1])
                nama = " ".join(tokens[:-1]).strip()
            else:
                jml = 1
                nama = left_part
            parsed.append({
                "KEBUTUHAN KEGIATAN": nama or "Item Borongan",
                "JUMLAH": jml,
                "HARGA SATUAN": 0,
                "HARGA TOTAL": h_tot,
                "KETERANGAN": default_ket,
            })
            continue

        tokens = clean_ln.split()
        if len(tokens) >= 2:
            last_tok = tokens[-1]
            second_last = tokens[-2] if len(tokens) >= 2 else ""
            if last_tok in ("-", "0", "nihil") and second_last.isdigit():
                jml = int(second_last)
                nama = " ".join(tokens[:-2]).strip()
                parsed.append({
                    "KEBUTUHAN KEGIATAN": nama or "Fasilitas",
                    "JUMLAH": jml,
                    "HARGA SATUAN": 0,
                    "HARGA TOTAL": 0,
                    "KETERANGAN": default_ket,
                })
                continue
            elif last_tok.isdigit() and int(last_tok) < 500:
                jml = int(last_tok)
                nama = " ".join(tokens[:-1]).strip()
                parsed.append({
                    "KEBUTUHAN KEGIATAN": nama or "Fasilitas",
                    "JUMLAH": jml,
                    "HARGA SATUAN": 0,
                    "HARGA TOTAL": 0,
                    "KETERANGAN": default_ket,
                })
                continue
            else:
                nominal = parse_num(last_tok)
                if second_last.isdigit() and int(second_last) < 1000 and nominal > 0:
                    jml = int(second_last)
                    nama = " ".join(tokens[:-2]).strip()
                    parsed.append({
                        "KEBUTUHAN KEGIATAN": nama or "Item Kegiatan",
                        "JUMLAH": jml,
                        "HARGA SATUAN": nominal,
                        "HARGA TOTAL": jml * nominal,
                        "KETERANGAN": default_ket,
                    })
                    continue

        parsed.append({
            "KEBUTUHAN KEGIATAN": clean_ln,
            "JUMLAH": 1,
            "HARGA SATUAN": 0,
            "HARGA TOTAL": 0,
            "KETERANGAN": default_ket,
        })

    return parsed


def build_paragraf_pembuka(nama_kegiatan: str, tema_kegiatan: str, konteks_hut: str) -> str:
    """Menyusun paragraf pembuka LPJ adaptif sesuai ketersediaan Tema dan Konteks HUT."""
    nama = (nama_kegiatan or "").strip() or "Kegiatan Dharma Wanita Persatuan KPPBC TMC Malang"
    tema = (tema_kegiatan or "").strip().strip('"')
    hut = (konteks_hut or "").strip()

    kalimat_1 = (
        "Puji syukur kehadirat Allah SWT yang telah melimpahkan rahmat dan rizki-Nya "
        f"kepada kita semua, yang mana telah terlaksana {nama}."
    )
    if tema and hut:
        kalimat_2 = f' Kegiatan ini mengusung tema "{tema}" serta {hut}, yang dilaksanakan pada :'
    elif tema:
        kalimat_2 = f' Kegiatan ini mengusung tema "{tema}", yang dilaksanakan pada :'
    elif hut:
        kalimat_2 = f" Kegiatan ini dilaksanakan dalam rangka {hut}, pada :"
    else:
        kalimat_2 = " Kegiatan tersebut dilaksanakan pada :"
    return kalimat_1 + kalimat_2


# ==============================================================================
# BAGIAN 3: KONEKSI GOOGLE SHEETS API (3 WORKSHEETS) & OFFLINE RESILIENCE
# ==============================================================================

@st.cache_resource(show_spinner=False)
def get_gspread_client(creds_json_str: str = ""):
    """Membuka koneksi gspread dengan dukungan st.secrets atau JSON Service Account."""
    if not HAS_GSPREAD:
        return None, "Library gspread belum terpasang di environment ini."

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = None

    if creds_json_str and creds_json_str.strip():
        try:
            creds_dict = json.loads(creds_json_str)
        except Exception as e:
            return None, f"Format JSON Service Account tidak valid: {e}"
    else:
        try:
            if "gcp_service_account" in st.secrets:
                creds_dict = dict(st.secrets["gcp_service_account"])
        except Exception:
            creds_dict = None

    if not creds_dict:
        return None, "Mode Lokal Aktif (Kredensial Cloud belum dihubungkan)"

    try:
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client, "Terhubung ke Google Sheets Cloud"
    except Exception as e:
        return None, f"Gagal otentikasi Google Sheets: {e}"


def ensure_worksheets_ready(sh):
    """Memastikan 3 worksheet (DATABASE_RAB, DATABASE_LPJ, REKAP_ITEM_DANA) siap digunakan."""
    existing_titles = [ws.title for ws in sh.worksheets()]

    if SHEET_DB_RAB not in existing_titles:
        ws_rab = sh.add_worksheet(title=SHEET_DB_RAB, rows=500, cols=len(COLS_DB_RAB))
        ws_rab.batch_update([{"range": "A1:J1", "values": [COLS_DB_RAB]}], value_input_option="USER_ENTERED")
    else:
        ws_rab = sh.worksheet(SHEET_DB_RAB)

    if SHEET_DB_LPJ not in existing_titles:
        ws_lpj = sh.add_worksheet(title=SHEET_DB_LPJ, rows=500, cols=len(COLS_DB_LPJ))
        ws_lpj.batch_update([{"range": "A1:O1", "values": [COLS_DB_LPJ]}], value_input_option="USER_ENTERED")
    else:
        ws_lpj = sh.worksheet(SHEET_DB_LPJ)

    if SHEET_REKAP_DANA not in existing_titles:
        ws_rk = sh.add_worksheet(title=SHEET_REKAP_DANA, rows=2500, cols=len(COLS_REKAP_DANA))
        ws_rk.batch_update([{"range": "A1:J1", "values": [COLS_REKAP_DANA]}], value_input_option="USER_ENTERED")
    else:
        ws_rk = sh.worksheet(SHEET_REKAP_DANA)

    return ws_rab, ws_lpj, ws_rk


@st.cache_data(ttl=180, show_spinner=False)
def fetch_sheet_bundle_cached(spreadsheet_id: str, creds_json_str: str = "") -> dict:
    """Membaca DATABASE_RAB, DATABASE_LPJ, dan REKAP_ITEM_DANA sekaligus dengan 1 panggilan API."""
    client, status_msg = get_gspread_client(creds_json_str)
    if client is None:
        return {"ok": False, "msg": status_msg, "rab_rows": [], "lpj_rows": [], "rekap_rows": []}

    try:
        sh = client.open_by_key(spreadsheet_id)
        ensure_worksheets_ready(sh)
        batch_res = sh.values_batch_get([
            f"{SHEET_DB_RAB}!A1:J500",
            f"{SHEET_DB_LPJ}!A1:O500",
            f"{SHEET_REKAP_DANA}!A1:J2500",
        ])
        value_ranges = batch_res.get("valueRanges", [])

        rab_vals = value_ranges[0].get("values", []) if len(value_ranges) > 0 else []
        lpj_vals = value_ranges[1].get("values", []) if len(value_ranges) > 1 else []
        rk_vals = value_ranges[2].get("values", []) if len(value_ranges) > 2 else []

        def _to_records(vals, expected_cols):
            if not vals or len(vals) < 2:
                return []
            headers = vals[0]
            out = []
            for r in vals[1:]:
                if not any(str(c).strip() for c in r):
                    continue
                row_dict = {}
                for i, col_name in enumerate(expected_cols):
                    col_idx = headers.index(col_name) if col_name in headers else i
                    row_dict[col_name] = r[col_idx] if col_idx < len(r) else ""
                out.append(row_dict)
            return out

        return {
            "ok": True,
            "msg": "Cloud Google Sheets Aktif",
            "rab_rows": _to_records(rab_vals, COLS_DB_RAB),
            "lpj_rows": _to_records(lpj_vals, COLS_DB_LPJ),
            "rekap_rows": _to_records(rk_vals, COLS_REKAP_DANA),
        }
    except Exception as e:
        return {"ok": False, "msg": f"Gagal membaca GSheet: {e}", "rab_rows": [], "lpj_rows": [], "rekap_rows": []}


def get_sheet_bundle(spreadsheet_id: str, creds_json_str: str = "") -> dict:
    """Mengambil data Google Sheets dengan fallback ke _backup_sheet_bundle & session_state lokal."""
    res = fetch_sheet_bundle_cached(spreadsheet_id, creds_json_str)
    if res.get("ok"):
        st.session_state["_backup_sheet_bundle"] = res
        return res

    if "_backup_sheet_bundle" in st.session_state:
        backup = dict(st.session_state["_backup_sheet_bundle"])
        backup["msg"] = backup.get("msg", "") + " (Cadangan Memori)"
        return backup

    return {
        "ok": False,
        "msg": res.get("msg", "Mode Offline Lokal Aktif"),
        "rab_rows": st.session_state.get("local_db_rab", []),
        "lpj_rows": st.session_state.get("local_db_lpj", []),
        "rekap_rows": st.session_state.get("local_rekap_dana", []),
    }


def save_rab_to_storage(rab_payload: dict, spreadsheet_id: str, creds_json_str: str = "") -> tuple:
    """Menyimpan dokumen RAB ke memori lokal dan Google Sheets (DATABASE_RAB & REKAP_ITEM_DANA)."""
    id_rab = rab_payload.get("id_rab") or ("RAB-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
    rab_payload["id_rab"] = id_rab
    waktu_simpan = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df_items = normalize_items_df(rab_payload.get("items_rab", []), rab_payload.get("ket_tabel_rab", ""))
    total_rab = int(df_items["HARGA TOTAL"].sum()) if not df_items.empty else 0

    serializable = dict(rab_payload)
    serializable["items_rab"] = df_items.to_dict(orient="records")

    rab_row_dict = {
        "ID_RAB": id_rab,
        "WAKTU_SIMPAN": waktu_simpan,
        "NAMA_KEGIATAN": rab_payload.get("nama_kegiatan_rab", ""),
        "LOKASI_ACARA": rab_payload.get("lokasi_acara_rab", ""),
        "SUMBER_DANA": rab_payload.get("sumber_dana_rab", ""),
        "KET_TABEL_DANA": rab_payload.get("ket_tabel_rab", ""),
        "TOTAL_RAB": total_rab,
        "CATATAN_NOTE": rab_payload.get("catatan_note_rab", ""),
        "STATUS_KONVERSI": rab_payload.get("status_konversi", "BELUM DIKONVERSI"),
        "PAYLOAD_JSON": json.dumps(serializable, ensure_ascii=False),
    }

    rekap_dicts = []
    for _, r in df_items.iterrows():
        rekap_dicts.append({
            "ID_DOKUMEN": id_rab,
            "JENIS_DOK": "RAB",
            "TGL_ACARA": rab_payload.get("lokasi_acara_rab", ""),
            "NAMA_KEGIATAN": rab_payload.get("nama_kegiatan_rab", ""),
            "NO_ITEM": int(r["NO"]),
            "KEBUTUHAN_KEGIATAN": str(r["KEBUTUHAN KEGIATAN"]),
            "JUMLAH": int(r["JUMLAH"]),
            "HARGA_SATUAN": int(r["HARGA SATUAN"]),
            "HARGA_TOTAL": int(r["HARGA TOTAL"]),
            "KETERANGAN": str(r["KETERANGAN"]),
        })

    local_rab = [row for row in st.session_state.get("local_db_rab", []) if row.get("ID_RAB") != id_rab]
    local_rab.insert(0, rab_row_dict)
    st.session_state["local_db_rab"] = local_rab

    local_rk = [row for row in st.session_state.get("local_rekap_dana", []) if row.get("ID_DOKUMEN") != id_rab]
    local_rk.extend(rekap_dicts)
    st.session_state["local_rekap_dana"] = local_rk

    client, status_msg = get_gspread_client(creds_json_str)
    if client is None:
        return True, f"RAB ({id_rab}) tersimpan di Memori Lokal. {status_msg}"

    try:
        sh = client.open_by_key(spreadsheet_id)
        ws_rab, _, ws_rk = ensure_worksheets_ready(sh)

        existing_vals = ws_rab.get_all_values()
        new_matrix = [COLS_DB_RAB]
        updated = False
        for r in existing_vals[1:]:
            if not any(str(c).strip() for c in r):
                continue
            if len(r) > 0 and r[0] == id_rab:
                new_matrix.append([rab_row_dict[c] for c in COLS_DB_RAB])
                updated = True
            else:
                padded = list(r) + [""] * max(0, len(COLS_DB_RAB) - len(r))
                new_matrix.append(padded[:len(COLS_DB_RAB)])
        if not updated:
            new_matrix.append([rab_row_dict[c] for c in COLS_DB_RAB])

        ws_rab.clear()
        ws_rab.batch_update(
            [{"range": f"A1:J{len(new_matrix)}", "values": new_matrix}],
            value_input_option="USER_ENTERED",
        )

        existing_rk = ws_rk.get_all_values()
        new_rk = [COLS_REKAP_DANA]
        for r in existing_rk[1:]:
            if not any(str(c).strip() for c in r):
                continue
            if len(r) > 0 and r[0] != id_rab:
                padded = list(r) + [""] * max(0, len(COLS_REKAP_DANA) - len(r))
                new_rk.append(padded[:len(COLS_REKAP_DANA)])
        for d in rekap_dicts:
            new_rk.append([d[c] for c in COLS_REKAP_DANA])

        ws_rk.clear()
        ws_rk.batch_update(
            [{"range": f"A1:J{len(new_rk)}", "values": new_rk}],
            value_input_option="USER_ENTERED",
        )

        fetch_sheet_bundle_cached.clear()
        return True, f"RAB ({id_rab}) berhasil disimpan ke Google Sheets ({SHEET_DB_RAB} & {SHEET_REKAP_DANA})!"
    except Exception as e:
        return True, f"RAB ({id_rab}) tersimpan di Lokal, namun gagal sinkron ke Cloud: {e}"


def save_lpj_to_storage(lpj_payload: dict, spreadsheet_id: str, creds_json_str: str = "") -> tuple:
    """Menyimpan dokumen LPJ ke memori lokal dan Google Sheets (DATABASE_LPJ & REKAP_ITEM_DANA)."""
    id_lpj = lpj_payload.get("id_lpj") or ("LPJ-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
    lpj_payload["id_lpj"] = id_lpj
    waktu_simpan = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    tgl_lap_str = str(lpj_payload.get("tgl_laporan", date.today()))
    tgl_acara_str = lpj_payload.get("tgl_acara_manual", "").strip() or format_hari_tanggal_indo(lpj_payload.get("tgl_acara", date.today()))

    df_items = normalize_items_df(lpj_payload.get("items_dana", []), lpj_payload.get("ket_tabel_dana", ""))
    total_dana = int(df_items["HARGA TOTAL"].sum()) if not df_items.empty else 0

    serializable = dict(lpj_payload)
    if isinstance(serializable.get("tgl_laporan"), (date, datetime)):
        serializable["tgl_laporan"] = serializable["tgl_laporan"].strftime("%Y-%m-%d")
    if isinstance(serializable.get("tgl_acara"), (date, datetime)):
        serializable["tgl_acara"] = serializable["tgl_acara"].strftime("%Y-%m-%d")
    serializable["items_dana"] = df_items.to_dict(orient="records")

    lpj_row_dict = {
        "ID_LPJ": id_lpj,
        "ID_RAB_REF": lpj_payload.get("id_rab_ref", ""),
        "WAKTU_SIMPAN": waktu_simpan,
        "TGL_LAPORAN": tgl_lap_str,
        "NAMA_KEGIATAN": lpj_payload.get("nama_kegiatan", ""),
        "TEMA_KEGIATAN": lpj_payload.get("tema_kegiatan", ""),
        "KONTEKS_HUT": lpj_payload.get("konteks_hut", ""),
        "TGL_ACARA": tgl_acara_str,
        "WAKTU_ACARA": lpj_payload.get("waktu_acara", ""),
        "TEMPAT_ACARA": lpj_payload.get("tempat_acara", ""),
        "SUMBER_DANA": lpj_payload.get("sumber_dana", ""),
        "KET_TABEL_DANA": lpj_payload.get("ket_tabel_dana", ""),
        "TOTAL_DANA": total_dana,
        "PENANDATANGAN": lpj_payload.get("nama_penandatangan", ""),
        "PAYLOAD_JSON": json.dumps(serializable, ensure_ascii=False),
    }

    rekap_dicts = []
    for _, r in df_items.iterrows():
        rekap_dicts.append({
            "ID_DOKUMEN": id_lpj,
            "JENIS_DOK": "LPJ",
            "TGL_ACARA": tgl_acara_str,
            "NAMA_KEGIATAN": lpj_payload.get("nama_kegiatan", ""),
            "NO_ITEM": int(r["NO"]),
            "KEBUTUHAN_KEGIATAN": str(r["KEBUTUHAN KEGIATAN"]),
            "JUMLAH": int(r["JUMLAH"]),
            "HARGA_SATUAN": int(r["HARGA SATUAN"]),
            "HARGA_TOTAL": int(r["HARGA TOTAL"]),
            "KETERANGAN": str(r["KETERANGAN"]),
        })

    local_lpj = [row for row in st.session_state.get("local_db_lpj", []) if row.get("ID_LPJ") != id_lpj]
    local_lpj.insert(0, lpj_row_dict)
    st.session_state["local_db_lpj"] = local_lpj

    local_rk = [row for row in st.session_state.get("local_rekap_dana", []) if row.get("ID_DOKUMEN") != id_lpj]
    local_rk.extend(rekap_dicts)
    st.session_state["local_rekap_dana"] = local_rk

    client, status_msg = get_gspread_client(creds_json_str)
    if client is None:
        return True, f"LPJ ({id_lpj}) tersimpan di Memori Lokal. {status_msg}"

    try:
        sh = client.open_by_key(spreadsheet_id)
        _, ws_lpj, ws_rk = ensure_worksheets_ready(sh)

        existing_vals = ws_lpj.get_all_values()
        new_matrix = [COLS_DB_LPJ]
        updated = False
        for r in existing_vals[1:]:
            if not any(str(c).strip() for c in r):
                continue
            if len(r) > 0 and r[0] == id_lpj:
                new_matrix.append([lpj_row_dict[c] for c in COLS_DB_LPJ])
                updated = True
            else:
                padded = list(r) + [""] * max(0, len(COLS_DB_LPJ) - len(r))
                new_matrix.append(padded[:len(COLS_DB_LPJ)])
        if not updated:
            new_matrix.append([lpj_row_dict[c] for c in COLS_DB_LPJ])

        ws_lpj.clear()
        ws_lpj.batch_update(
            [{"range": f"A1:O{len(new_matrix)}", "values": new_matrix}],
            value_input_option="USER_ENTERED",
        )

        existing_rk = ws_rk.get_all_values()
        new_rk = [COLS_REKAP_DANA]
        for r in existing_rk[1:]:
            if not any(str(c).strip() for c in r):
                continue
            if len(r) > 0 and r[0] != id_lpj:
                padded = list(r) + [""] * max(0, len(COLS_REKAP_DANA) - len(r))
                new_rk.append(padded[:len(COLS_REKAP_DANA)])
        for d in rekap_dicts:
            new_rk.append([d[c] for c in COLS_REKAP_DANA])

        ws_rk.clear()
        ws_rk.batch_update(
            [{"range": f"A1:J{len(new_rk)}", "values": new_rk}],
            value_input_option="USER_ENTERED",
        )

        fetch_sheet_bundle_cached.clear()
        return True, f"LPJ ({id_lpj}) berhasil disimpan ke Google Sheets ({SHEET_DB_LPJ} & {SHEET_REKAP_DANA})!"
    except Exception as e:
        return True, f"LPJ ({id_lpj}) tersimpan di Lokal, namun gagal sinkron ke Cloud: {e}"


# ==============================================================================
# BAGIAN 4: ENGINE DOKUMEN WORD (RAB & LPJ) + LIBREOFFICE HEADLESS PDF CACHED
# ==============================================================================

def find_template_file(filename: str) -> str:
    """Mencari file template .docx di direktori kerja atau subdirektori."""
    candidates = [
        filename,
        os.path.join(os.getcwd(), filename),
    ]
    for root_dir, _, files in os.walk(os.getcwd()):
        if filename in files:
            candidates.append(os.path.join(root_dir, filename))
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return ""


def get_default_assets_bytes() -> dict:
    """Mengembalikan byte gambar Logo Kiri, Logo Kanan, dan Tanda Tangan Asli Ny. Rini Johan Pandores."""
    logo1_bytes = base64.b64decode(EMBEDDED_LOGO1_B64)
    logo2_bytes = base64.b64decode(EMBEDDED_LOGO2_B64)
    sig_bytes = base64.b64decode(EMBEDDED_SIG_B64)

    tpl_bytes = st.session_state.get("uploaded_template_bytes")
    tpl_path = find_template_file(TEMPLATE_LPJ_FILENAME)

    try:
        if tpl_bytes:
            doc_tmp = docx.Document(io.BytesIO(tpl_bytes))
        elif tpl_path:
            doc_tmp = docx.Document(tpl_path)
        else:
            doc_tmp = None

        if doc_tmp is not None:
            for _, part in doc_tmp.part.related_parts.items():
                if "image" in part.content_type:
                    sig_bytes = part.blob
    except Exception:
        pass

    return {
        "logo1": logo1_bytes,
        "logo2": logo2_bytes,
        "sig": sig_bytes,
    }


def _set_cell_margins(cell, top=42, bottom=42, left=65, right=65):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m_name, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m_name}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_cell_shading(cell, hex_color: str):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def _set_table_borders(table, color="505050", sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def _build_fallback_header(section, logo1_bytes: bytes, logo2_bytes: bytes):
    """Membangun Kop Surat Resmi DWP KPPBC TMC Malang secara terprogram jika file template tidak ditemukan."""
    header = section.header
    header.is_linked_to_previous = False
    for p in list(header.paragraphs):
        p._element.getparent().remove(p._element)

    htbl = header.add_table(rows=1, cols=3, width=Cm(19.5))
    htbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(2.5), Cm(14.2), Cm(2.8)]
    for i, w in enumerate(widths):
        htbl.columns[i].width = w
        htbl.rows[0].cells[i].width = w
        htbl.rows[0].cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p_left = htbl.rows[0].cells[0].paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left.add_run().add_picture(io.BytesIO(logo1_bytes), width=Cm(1.95))

    c_mid = htbl.rows[0].cells[1]
    p1 = c_mid.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_after = Pt(1)
    r1 = p1.add_run("Dharma Wanita Persatuan")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(16)
    r1.font.bold = True
    r1.font.italic = True

    p2 = c_mid.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(1)
    r2 = p2.add_run(
        "DIREKTORAT JENDERAL BEA DAN CUKAI KANTOR WILAYAH DJBC JAWA TIMUR II" + NL +
        "KANTOR PENGAWASAN DAN PELAYANAN BEA DAN CUKAI TIPE MADYA CUKAI MALANG"
    )
    r2.font.name = "Arial"
    r2.font.size = Pt(9)
    r2.font.bold = True

    p3 = c_mid.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(0)
    r3 = p3.add_run(
        "JALAN SURABAYA NOMOR 2, MALANG 65145" + NL +
        "TELEPON (0341) 551628; FAKSIMILE (0341)554193; LAMAN WWW.BEACUKAIMALANG.COM" + NL +
        "PUSAT KONTAK LAYANAN 1500225; SUREL INFO@CUSTOMS.GO.ID"
    )
    r3.font.name = "Calibri"
    r3.font.size = Pt(6.5)

    p_right = htbl.rows[0].cells[2].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right.add_run().add_picture(io.BytesIO(logo2_bytes), width=Cm(2.25))

    p_line = header.add_paragraph()
    p_line.paragraph_format.space_before = Pt(2)
    p_line.paragraph_format.space_after = Pt(4)
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        '<w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/>'
        "</w:pBdr>"
    )
    p_line._p.get_or_add_pPr().append(pBdr)


def _prepare_cloned_document():
    """Membuka dokumen dengan Kop Surat Resmi DWP KPPBC TMC Malang yang di-clone 100% utuh."""
    assets = get_default_assets_bytes()
    tpl_bytes = st.session_state.get("uploaded_template_bytes")
    tpl_path = find_template_file(TEMPLATE_LPJ_FILENAME) or find_template_file(TEMPLATE_RAB_FILENAME)

    used_cloned = False
    if tpl_bytes:
        doc = docx.Document(io.BytesIO(tpl_bytes))
        used_cloned = True
    elif tpl_path:
        doc = docx.Document(tpl_path)
        used_cloned = True
    else:
        doc = docx.Document()

    if used_cloned:
        hdr_ref = None
        for p in doc.paragraphs:
            pPr = p._p.find(qn("w:pPr"))
            if pPr is not None:
                sPr = pPr.find(qn("w:sectPr"))
                if sPr is not None:
                    href = sPr.find(qn("w:headerReference"))
                    if href is not None:
                        hdr_ref = href

        body_el = doc._body._body
        final_sectPr = body_el.find(qn("w:sectPr"))
        if hdr_ref is not None and final_sectPr is not None:
            if final_sectPr.find(qn("w:headerReference")) is None:
                final_sectPr.insert(0, hdr_ref)

        for child in list(body_el):
            if child.tag != qn("w:sectPr"):
                body_el.remove(child)
    else:
        sec = doc.sections[0]
        sec.page_width = Cm(21.59)
        sec.page_height = Cm(27.94)
        sec.top_margin = Cm(3.85)
        sec.bottom_margin = Cm(1.20)
        sec.left_margin = Cm(0.88)
        sec.right_margin = Cm(0.50)
        sec.header_distance = Cm(1.28)
        _build_fallback_header(sec, assets["logo1"], assets["logo2"])

    for sec in doc.sections:
        if sec.bottom_margin < Cm(1.0):
            sec.bottom_margin = Cm(1.0)

    return doc, assets


def _build_6col_financial_table(doc, df_items: pd.DataFrame, merge_ket: bool, ket_global: str):
    """Membangun Tabel 6 Kolom Presisi (Anti-Potong & Anti-Lompat Baris) untuk RAB maupun LPJ."""
    n_items = len(df_items)
    total_rows = n_items + 3
    col_widths_cm = [1.05, 5.80, 1.85, 2.95, 3.10, 4.45]

    table = doc.add_table(rows=total_rows, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_borders(table, color="505050", sz="4")

    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblPr.append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="400" w:type="dxa"/>'))

    for col_idx, w_cm in enumerate(col_widths_cm):
        table.columns[col_idx].width = Cm(w_cm)

    def _format_cell_text(cell, text: str, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5, no_wrap=False):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_margins(cell, top=42, bottom=42, left=65, right=65)
        if no_wrap:
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(OxmlElement("w:noWrap"))
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        for r_old in list(p.runs):
            p._p.remove(r_old._r)
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(size_pt)
        r.font.bold = bold
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    r0 = table.rows[0]
    r1 = table.rows[1]
    for h_row in (r0, r1):
        trPr = h_row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:tblHeader"))
        trPr.append(OxmlElement("w:cantSplit"))

    r0.cells[3].merge(r0.cells[4])
    for c_idx in (0, 1, 2, 5):
        r0.cells[c_idx].merge(r1.cells[c_idx])

    _format_cell_text(r0.cells[0], "NO", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5, no_wrap=True)
    _format_cell_text(r0.cells[1], "KEBUTUHAN KEGIATAN", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5)
    _format_cell_text(r0.cells[2], "JUMLAH", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5, no_wrap=True)
    _format_cell_text(r0.cells[3], "DANA", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5)
    _format_cell_text(r0.cells[5], "KETERANGAN", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5)

    _format_cell_text(r1.cells[3], "HARGA SATUAN", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.0, no_wrap=True)
    _format_cell_text(r1.cells[4], "HARGA TOTAL", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.0, no_wrap=True)

    for h_r in (r0, r1):
        for c_idx, w_cm in enumerate(col_widths_cm):
            h_r.cells[c_idx].width = Cm(w_cm)
            _set_cell_shading(h_r.cells[c_idx], "F8FAFC")

    for idx_df, row_data in df_items.iterrows():
        row_pos = 2 + idx_df
        tr_row = table.rows[row_pos]
        tr_row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))

        for c_idx, w_cm in enumerate(col_widths_cm):
            tr_row.cells[c_idx].width = Cm(w_cm)

        no_str = str(int(row_data["NO"]))
        keb_str = str(row_data["KEBUTUHAN KEGIATAN"])
        jml_str = str(int(row_data["JUMLAH"])) if int(row_data["JUMLAH"]) > 0 else "-"
        hsat_str = fmt_rp_cell(row_data["HARGA SATUAN"], is_dash_if_zero=True)
        htot_str = fmt_rp_cell(row_data["HARGA TOTAL"], is_dash_if_zero=True)
        ket_item = str(row_data.get("KETERANGAN", ket_global) or "")

        _format_cell_text(tr_row.cells[0], no_str, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5)
        _format_cell_text(tr_row.cells[1], keb_str, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, size_pt=10.5)
        _format_cell_text(tr_row.cells[2], jml_str, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5)
        _format_cell_text(tr_row.cells[3], hsat_str, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=False, size_pt=10.5, no_wrap=True)
        _format_cell_text(tr_row.cells[4], htot_str, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=False, size_pt=10.5, no_wrap=True)

        if not merge_ket:
            _format_cell_text(tr_row.cells[5], ket_item, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.0)

    if merge_ket and n_items > 0:
        first_ket_cell = table.rows[2].cells[5]
        if n_items > 1:
            last_ket_cell = table.rows[2 + n_items - 1].cells[5]
            first_ket_cell.merge(last_ket_cell)
        _format_cell_text(first_ket_cell, ket_global, align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5)

    r_tot = table.rows[total_rows - 1]
    r_tot._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    for c_idx, w_cm in enumerate(col_widths_cm):
        r_tot.cells[c_idx].width = Cm(w_cm)

    r_tot.cells[0].merge(r_tot.cells[3])
    total_keseluruhan = int(df_items["HARGA TOTAL"].sum()) if not df_items.empty else 0

    _format_cell_text(r_tot.cells[0], "TOTAL", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=10.5)
    _format_cell_text(r_tot.cells[4], fmt_rp_cell(total_keseluruhan, is_dash_if_zero=False), align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, size_pt=10.5, no_wrap=True)
    _format_cell_text(r_tot.cells[5], "", align=WD_ALIGN_PARAGRAPH.CENTER, bold=False, size_pt=10.5)


def generate_rab_docx_bytes(rab_payload: dict) -> bytes:
    """
    Mesin pembuat dokumen Word (.docx) khusus RAB sesuai pakem resmi RAB DWP GABUNGAN DI JEMBER.docx:
    - Kop Surat Resmi DWP KPPBC TMC Malang
    - Judul 3 Baris Terpusat (Bold, Kapital)
    - Kalimat Pengantar Sumber Dana
    - Tabel 6 Kolom Rincian Anggaran (Anti-Wrap & Anti-Potong)
    - Catatan Kaki (*Note :) dengan daftar ber-bullet
    """
    doc, _ = _prepare_cloned_document()

    def _add_p(text="", align=WD_ALIGN_PARAGRAPH.LEFT, left_cm=0.73, right_cm=0.80, first_line_cm=0.0, before_pt=0, after_pt=3, bold=False, italic=False, size_pt=11.0):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.left_indent = Cm(left_cm)
        pf.right_indent = Cm(right_cm)
        pf.first_line_indent = Cm(first_line_cm)
        pf.space_before = Pt(before_pt)
        pf.space_after = Pt(after_pt)
        pf.line_spacing = 1.15
        pf.keep_together = True
        if text:
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(size_pt)
            r.font.bold = bold
            r.font.italic = italic
            r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    # 1. Judul 3 Baris Terpusat (Bold, Kapital)
    b1 = (rab_payload.get("baris_1_judul") or "RANCANGAN ANGGARAN BIAYA").strip().upper()
    b2 = (rab_payload.get("nama_kegiatan_rab") or "PELAKSANAAN PERTEMUAN RUTIN PAGUYUBAN DWP").strip().upper()
    b3 = (rab_payload.get("lokasi_acara_rab") or "").strip().upper()
    if b3 and not b3.startswith("("):
        b3 = f"({b3})"

    p_t1 = _add_p(b1, align=WD_ALIGN_PARAGRAPH.CENTER, left_cm=0.73, right_cm=0.80, before_pt=10, after_pt=1, bold=True, size_pt=11.0)
    p_t1.paragraph_format.keep_with_next = True

    p_t2 = _add_p(b2, align=WD_ALIGN_PARAGRAPH.CENTER, left_cm=0.73, right_cm=0.80, before_pt=0, after_pt=1, bold=True, size_pt=11.0)
    p_t2.paragraph_format.keep_with_next = True

    if b3:
        p_t3 = _add_p(b3, align=WD_ALIGN_PARAGRAPH.CENTER, left_cm=0.73, right_cm=0.80, before_pt=0, after_pt=12, bold=True, size_pt=11.0)
        p_t3.paragraph_format.keep_with_next = True

    # 2. Kalimat Pengantar Sumber Dana
    sumber = (rab_payload.get("sumber_dana_rab") or "Kas Dharma Wanita KPPBC TMC Malang").strip()
    p_pengantar = _add_p(
        f"Dana untuk kegiatan ini berasal dari {sumber}. Dengan rincian dana sebagai berikut :",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        left_cm=0.73,
        right_cm=0.80,
        first_line_cm=1.25,
        before_pt=4,
        after_pt=8,
        size_pt=11.0,
    )
    p_pengantar.paragraph_format.keep_with_next = True

    # 3. Tabel 6 Kolom Rincian Anggaran
    df_items = normalize_items_df(rab_payload.get("items_rab", []), rab_payload.get("ket_tabel_rab", ""))
    _build_6col_financial_table(
        doc,
        df_items=df_items,
        merge_ket=bool(rab_payload.get("merge_ket_rab", False)),
        ket_global=str(rab_payload.get("ket_tabel_rab", "") or ""),
    )

    # 4. Catatan Kaki (*Note :) dengan Bullet
    note_raw = (rab_payload.get("catatan_note_rab") or "").strip()
    note_lines = [ln.strip() for ln in note_raw.split(NL) if ln.strip()]
    if note_lines:
        p_note_hdr = _add_p(
            "*Note :",
            align=WD_ALIGN_PARAGRAPH.LEFT,
            left_cm=0.73,
            right_cm=0.80,
            before_pt=10,
            after_pt=3,
            italic=True,
            size_pt=10.0,
        )
        p_note_hdr.paragraph_format.keep_with_next = True

        for ln in note_lines:
            clean_ln = ln.lstrip("•-* ").strip()
            p_b = doc.add_paragraph()
            p_b.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf_b = p_b.paragraph_format
            pf_b.left_indent = Cm(1.55)
            pf_b.right_indent = Cm(0.80)
            pf_b.first_line_indent = Cm(-0.50)
            pf_b.space_before = Pt(1)
            pf_b.space_after = Pt(2)
            pf_b.line_spacing = 1.12
            pf_b.tab_stops.add_tab_stop(Cm(1.55), WD_TAB_ALIGNMENT.LEFT)
            r_b = p_b.add_run(BULLET_CHAR + TAB + clean_ln)
            r_b.font.name = "Arial"
            r_b.font.size = Pt(10.0)

    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue()


def generate_lpj_docx_bytes(payload: dict, custom_sig_bytes: bytes = None, attachment_images: list = None) -> bytes:
    """
    Mesin utama pembuat dokumen Word (.docx) LPJ sesuai pakem resmi LPJ_DWP GABUNGAN DI BLITAR_2026-1.docx.
    """
    doc, assets = _prepare_cloned_document()
    sig_to_use = custom_sig_bytes if custom_sig_bytes else assets["sig"]

    def _add_styled_p(text="", align=WD_ALIGN_PARAGRAPH.LEFT, left_cm=0.73, right_cm=0.80, first_line_cm=0.0, space_before_pt=0, space_after_pt=3, italic=False, bold=False, size_pt=11.0):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.left_indent = Cm(left_cm)
        pf.right_indent = Cm(right_cm)
        pf.first_line_indent = Cm(first_line_cm)
        pf.space_before = Pt(space_before_pt)
        pf.space_after = Pt(space_after_pt)
        pf.line_spacing = 1.15
        pf.keep_together = True
        if text:
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(size_pt)
            r.font.italic = italic
            r.font.bold = bold
            r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    # 1. Tanggal Laporan (Rata Kanan)
    tgl_laporan_indo = format_hari_tanggal_indo(payload.get("tgl_laporan", date.today()))
    _add_styled_p(
        tgl_laporan_indo,
        align=WD_ALIGN_PARAGRAPH.RIGHT,
        left_cm=0.73,
        right_cm=1.32,
        space_before_pt=6,
        space_after_pt=10,
        size_pt=11.0,
    )

    # 2. Salam Pembuka
    _add_styled_p(
        "Assalamu’alaikum Wr. Wb.",
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=0.73,
        right_cm=0.80,
        space_before_pt=2,
        space_after_pt=8,
        italic=True,
        size_pt=11.0,
    )

    # 3. Paragraf Pembuka Adaptif
    teks_pembuka = build_paragraf_pembuka(
        payload.get("nama_kegiatan", ""),
        payload.get("tema_kegiatan", ""),
        payload.get("konteks_hut", ""),
    )
    _add_styled_p(
        teks_pembuka,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        left_cm=0.73,
        right_cm=0.80,
        first_line_cm=1.27,
        space_before_pt=0,
        space_after_pt=6,
        size_pt=11.0,
    )

    # 4. Blok Waktu & Tempat (Menggunakan Tab-Stop Presisi)
    tgl_acara_indo = (payload.get("tgl_acara_manual") or "").strip()
    if not tgl_acara_indo:
        tgl_acara_indo = format_hari_tanggal_indo(payload.get("tgl_acara", date.today()))

    waktu_acara = (payload.get("waktu_acara") or "08.30 - selesai").strip()
    tempat_acara = (payload.get("tempat_acara") or "-").strip()

    for label_kiri, nilai_kanan in [
        ("Hari, Tanggal", tgl_acara_indo),
        ("Waktu", waktu_acara),
        ("Tempat", tempat_acara),
    ]:
        p_info = _add_styled_p(
            "",
            align=WD_ALIGN_PARAGRAPH.LEFT,
            left_cm=1.98,
            right_cm=0.80,
            space_before_pt=1,
            space_after_pt=2,
            size_pt=11.0,
        )
        p_info.paragraph_format.tab_stops.add_tab_stop(Cm(4.85), WD_TAB_ALIGNMENT.LEFT)
        r_lbl = p_info.add_run(label_kiri + TAB + ": " + nilai_kanan)
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(11.0)

    # 5. Rangkaian Kegiatan
    _add_styled_p(
        "Dengan rangkaian kegiatan sebagai berikut :",
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=1.98,
        right_cm=0.80,
        space_before_pt=6,
        space_after_pt=3,
        size_pt=11.0,
    )

    rangkaian_raw = (payload.get("rangkaian_kegiatan") or "").strip()
    rangkaian_list = [ln.strip() for ln in rangkaian_raw.split(NL) if ln.strip()]
    for idx_r, item_r in enumerate(rangkaian_list, start=1):
        clean_r = item_r
        if "." in clean_r[:4]:
            l_num, r_rest = clean_r.split(".", 1)
            if l_num.strip().isdigit():
                clean_r = r_rest.strip()

        p_r = doc.add_paragraph()
        p_r.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf_r = p_r.paragraph_format
        pf_r.left_indent = Cm(2.61)
        pf_r.right_indent = Cm(0.80)
        pf_r.first_line_indent = Cm(-0.63)
        pf_r.space_before = Pt(1)
        pf_r.space_after = Pt(2)
        pf_r.line_spacing = 1.12
        pf_r.tab_stops.add_tab_stop(Cm(2.61), WD_TAB_ALIGNMENT.LEFT)
        run_r = p_r.add_run(f"{idx_r}." + TAB + clean_r)
        run_r.font.name = "Arial"
        run_r.font.size = Pt(11.0)

    # 6. Pengantar Rincian Dana
    sumber_dana = (payload.get("sumber_dana") or "Kas Dharma Wanita KPPBC TMC Malang dan donatur").strip()
    _add_styled_p(
        f"Dana untuk kegiatan ini berasal dari {sumber_dana}. Dengan rincian dana sebagai berikut :",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        left_cm=0.73,
        right_cm=0.80,
        first_line_cm=1.27,
        space_before_pt=6,
        space_after_pt=6,
        size_pt=11.0,
    )

    # 7. Tabel 6 Kolom Rincian Dana
    df_items = normalize_items_df(payload.get("items_dana", []), payload.get("ket_tabel_dana", ""))
    _build_6col_financial_table(
        doc,
        df_items=df_items,
        merge_ket=bool(payload.get("merge_ket_vertikal", True)),
        ket_global=str(payload.get("ket_tabel_dana", "Kas DWP BC Malang & Donatur Anggota") or ""),
    )

    # 8. Paragraf Ucapan Terima Kasih Donatur
    if payload.get("tampilkan_ucapan_donatur", True):
        teks_donatur = (payload.get("teks_ucapan_donatur") or DEFAULT_UCAPAN_DONATUR).strip()
        if teks_donatur:
            p_don = _add_styled_p(
                teks_donatur,
                align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                left_cm=0.73,
                right_cm=0.80,
                first_line_cm=1.25,
                space_before_pt=10,
                space_after_pt=6,
                size_pt=11.0,
            )
            p_don.paragraph_format.keep_with_next = True

    # 9. Paragraf Penutup
    nama_keg = (payload.get("nama_kegiatan") or "Kegiatan Dharma Wanita Persatuan").strip()
    teks_penutup = (
        f"Demikian laporan pertanggung jawaban {nama_keg} ini kami sampaikan, "
        "atas perhatian dan bantuan yang telah diberikan kami ucapkan terima kasih."
    )
    _add_styled_p(
        teks_penutup,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        left_cm=0.73,
        right_cm=0.80,
        first_line_cm=1.25,
        space_before_pt=4,
        space_after_pt=6,
        size_pt=11.0,
    )

    # 10. Salam Penutup
    _add_styled_p(
        "Wassalamu'alaikum Wr. Wb.",
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=0.73,
        right_cm=0.80,
        space_before_pt=2,
        space_after_pt=6,
        italic=True,
        size_pt=11.0,
    )

    # 11. Catatan Instagram
    catatan_ig = (payload.get("catatan_ig") or "").strip()
    if catatan_ig:
        p_ig = _add_styled_p(
            "",
            align=WD_ALIGN_PARAGRAPH.LEFT,
            left_cm=0.73,
            right_cm=0.80,
            space_before_pt=4,
            space_after_pt=8,
            size_pt=10.0,
        )
        if "@dwpbcmalang" in catatan_ig:
            parts_ig = catatan_ig.split("@dwpbcmalang", 1)
            r_ig1 = p_ig.add_run(parts_ig[0])
            r_ig1.font.name = "Arial"
            r_ig1.font.size = Pt(10.0)
            r_ig1.font.italic = True
            r_ig2 = p_ig.add_run("@dwpbcmalang")
            r_ig2.font.name = "Arial"
            r_ig2.font.size = Pt(10.0)
            r_ig2.font.bold = True
            if parts_ig[1]:
                r_ig3 = p_ig.add_run(parts_ig[1])
                r_ig3.font.name = "Arial"
                r_ig3.font.size = Pt(10.0)
        else:
            r_ig = p_ig.add_run(catatan_ig)
            r_ig.font.name = "Arial"
            r_ig.font.size = Pt(10.0)
            r_ig.font.italic = True

    # 12. Blok Tanda Tangan
    jab_1 = (payload.get("jabatan_baris_1") or "Ketua Dharma Wanita Persatuan").strip()
    jab_2 = (payload.get("jabatan_baris_2") or "KPPBC Tipe Madya Cukai Malang").strip()
    nama_ttd = (payload.get("nama_penandatangan") or "Ny. Rini Johan Pandores").strip()

    p_jab = _add_styled_p(
        jab_1 + NL + jab_2,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=12.14,
        right_cm=0.80,
        space_before_pt=6,
        space_after_pt=2,
        size_pt=11.0,
    )
    p_jab.paragraph_format.keep_with_next = True

    p_sig_img = doc.add_paragraph()
    p_sig_img.paragraph_format.left_indent = Cm(12.14)
    p_sig_img.paragraph_format.space_before = Pt(2)
    p_sig_img.paragraph_format.space_after = Pt(2)
    p_sig_img.paragraph_format.keep_with_next = True

    if payload.get("tampilkan_ttd", True) and sig_to_use:
        try:
            r_sig = p_sig_img.add_run()
            r_sig.add_picture(io.BytesIO(sig_to_use), width=Cm(3.35))
        except Exception:
            p_sig_img.paragraph_format.space_before = Pt(18)
            p_sig_img.paragraph_format.space_after = Pt(18)
    else:
        p_sig_img.paragraph_format.space_before = Pt(20)
        p_sig_img.paragraph_format.space_after = Pt(20)

    _add_styled_p(
        nama_ttd,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        left_cm=12.14,
        right_cm=0.80,
        space_before_pt=2,
        space_after_pt=4,
        size_pt=11.0,
    )

    # 13. Lampiran Foto Dokumentasi / Nota Kwitansi (Auto-Grid 2x2)
    if attachment_images:
        doc.add_page_break()
        p_lamp = _add_styled_p(
            "LAMPIRAN DOKUMENTASI & BUKTI KWITANSI KEGIATAN",
            align=WD_ALIGN_PARAGRAPH.CENTER,
            left_cm=0.73,
            right_cm=0.80,
            space_before_pt=6,
            space_after_pt=10,
            bold=True,
            size_pt=12.0,
        )
        p_lamp.paragraph_format.keep_with_next = True

        n_imgs = len(attachment_images)
        grid_rows = (n_imgs + 1) // 2
        gtbl = doc.add_table(rows=grid_rows, cols=2)
        gtbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_borders(gtbl, color="CBD5E1", sz="4")

        for idx_img, img_item in enumerate(attachment_images):
            r_i = idx_img // 2
            c_i = idx_img % 2
            cell = gtbl.rows[r_i].cells[c_i]
            cell.width = Cm(9.2)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
            p_c = cell.paragraphs[0]
            p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                p_c.add_run().add_picture(io.BytesIO(img_item["bytes"]), width=Cm(8.2))
                p_cap = cell.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_cap = p_cap.add_run(img_item.get("caption", f"Dokumentasi {idx_img + 1}"))
                r_cap.font.name = "Arial"
                r_cap.font.size = Pt(9.5)
                r_cap.font.italic = True
            except Exception:
                p_c.add_run("[Gambar tidak dapat dimuat]")

    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue()


@st.cache_data(show_spinner=False)
def convert_docx_bytes_to_pdf_bytes(docx_bytes: bytes) -> tuple:
    """Mengonversi byte .docx menjadi byte .pdf menggunakan LibreOffice Headless (Cached Anti-Lag)."""
    lo_bin = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo_bin:
        return None, "LibreOffice Headless tidak ditemukan di server. Gunakan tombol Cetak PDF Browser di bawah."

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            in_docx = os.path.join(tmpdir, "DOKUMEN_DWP.docx")
            out_pdf = os.path.join(tmpdir, "DOKUMEN_DWP.pdf")
            with open(in_docx, "wb") as f:
                f.write(docx_bytes)

            cmd = [
                lo_bin,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                tmpdir,
                in_docx,
            ]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=45)
            if os.path.isfile(out_pdf) and os.path.getsize(out_pdf) > 0:
                with open(out_pdf, "rb") as f_pdf:
                    return f_pdf.read(), "PDF berhasil dibuat dengan LibreOffice Headless."
            err_txt = proc.stderr.decode("utf-8", errors="ignore")
            return None, f"Konversi PDF gagal: {err_txt}"
    except Exception as e:
        return None, f"Kendala konversi PDF: {e}"


# ==============================================================================
# BAGIAN 5: DIRECT DOWNLOAD BASE64 & LIVE PREVIEW KERTAS A4 (RAB & LPJ)
# ==============================================================================

def make_direct_download_html(file_bytes: bytes, filename: str, mime_type: str, label: str, bg_gradient: str) -> str:
    """Membuat tombol unduh langsung berbasis Base64 Data URI (Anti-Error 404 Streamlit)."""
    b64_data = base64.b64encode(file_bytes).decode("ascii")
    return (
        f'<a href="data:{mime_type};base64,{b64_data}" download="{filename}" '
        f'style="display:flex;align-items:center;justify-content:center;gap:8px;'
        f'width:100%;min-height:48px;padding:10px 16px;border-radius:10px;'
        f'background:{bg_gradient};color:#FFFFFF !important;'
        f'-webkit-text-fill-color:#FFFFFF !important;font-weight:700;font-size:14.5px;'
        f'text-decoration:none;box-shadow:0 4px 12px rgba(15,23,42,0.14);'
        f'border:1px solid rgba(255,255,255,0.25);text-align:center;">'
        f"{label}</a>"
    )


def _build_kop_html(logo1_b64: str, logo2_b64: str) -> str:
    return f"""
    <table class="kop-table">
        <tr>
            <td style="width:78px;text-align:center;">
                <img src="data:image/jpeg;base64,{logo1_b64}" style="width:68px;height:auto;"/>
            </td>
            <td style="text-align:center;">
                <div style="font-family:'Times New Roman',serif;font-size:20px;font-weight:700;font-style:italic;letter-spacing:0.2px;">
                    Dharma Wanita Persatuan
                </div>
                <div style="font-size:11px;font-weight:700;line-height:1.25;margin-top:2px;">
                    DIREKTORAT JENDERAL BEA DAN CUKAI KANTOR WILAYAH DJBC JAWA TIMUR II<br/>
                    KANTOR PENGAWASAN DAN PELAYANAN BEA DAN CUKAI TIPE MADYA CUKAI MALANG
                </div>
                <div style="font-size:8.5px;margin-top:3px;line-height:1.25;">
                    JALAN SURABAYA NOMOR 2, MALANG 65145<br/>
                    TELEPON (0341) 551628; FAKSIMILE (0341)554193; LAMAN <span style="color:#1D4ED8;text-decoration:underline;">WWW.BEACUKAIMALANG.COM</span><br/>
                    PUSAT KONTAK LAYANAN 1500225; SUREL <span style="color:#1D4ED8;text-decoration:underline;">INFO@CUSTOMS.GO.ID</span>
                </div>
            </td>
            <td style="width:88px;text-align:center;">
                <img src="data:image/png;base64,{logo2_b64}" style="width:78px;height:auto;"/>
            </td>
        </tr>
    </table>
    """


def build_a4_live_preview_html(doc_mode: str, payload: dict, custom_sig_bytes: bytes = None) -> str:
    """Membangun pratinjau visual HTML Kertas A4 untuk RAB maupun LPJ."""
    assets = get_default_assets_bytes()
    logo1_b64 = base64.b64encode(assets["logo1"]).decode("ascii")
    logo2_b64 = base64.b64encode(assets["logo2"]).decode("ascii")
    kop_html = _build_kop_html(logo1_b64, logo2_b64)

    if doc_mode == "RAB":
        b1 = (payload.get("baris_1_judul") or "RANCANGAN ANGGARAN BIAYA").strip().upper()
        b2 = (payload.get("nama_kegiatan_rab") or "").strip().upper()
        b3 = (payload.get("lokasi_acara_rab") or "").strip().upper()
        if b3 and not b3.startswith("("):
            b3 = f"({b3})"

        sumber = (payload.get("sumber_dana_rab") or "Kas Dharma Wanita KPPBC TMC Malang").strip()
        df_items = normalize_items_df(payload.get("items_rab", []), payload.get("ket_tabel_rab", ""))
        n_items = len(df_items)
        merge_ket = bool(payload.get("merge_ket_rab", False))
        ket_global = (payload.get("ket_tabel_rab") or "").strip()
        total_dana = int(df_items["HARGA TOTAL"].sum()) if not df_items.empty else 0

        rows_html = ""
        for idx_df, r in df_items.iterrows():
            no_str = str(int(r["NO"]))
            keb_str = str(r["KEBUTUHAN KEGIATAN"])
            jml_str = str(int(r["JUMLAH"])) if int(r["JUMLAH"]) > 0 else "-"
            hsat_str = fmt_rp_cell(r["HARGA SATUAN"], is_dash_if_zero=True)
            htot_str = fmt_rp_cell(r["HARGA TOTAL"], is_dash_if_zero=True)
            ket_item = str(r.get("KETERANGAN", ket_global) or "")

            if merge_ket:
                if idx_df == 0:
                    ket_td = f"<td rowspan='{n_items}' style='text-align:center;vertical-align:middle;padding:6px;'>{ket_global}</td>"
                else:
                    ket_td = ""
            else:
                ket_td = f"<td style='text-align:center;vertical-align:middle;padding:4px 6px;'>{ket_item}</td>"

            rows_html += (
                "<tr>"
                f"<td style='text-align:center;'>{no_str}</td>"
                f"<td style='text-align:left;'>{keb_str}</td>"
                f"<td style='text-align:center;'>{jml_str}</td>"
                f"<td style='text-align:right;white-space:nowrap;'>{hsat_str}</td>"
                f"<td style='text-align:right;white-space:nowrap;'>{htot_str}</td>"
                f"{ket_td}"
                "</tr>"
            )

        note_raw = (payload.get("catatan_note_rab") or "").strip()
        note_lines = [ln.strip() for ln in note_raw.split(NL) if ln.strip()]
        note_html = ""
        if note_lines:
            li_items = "".join(f"<li style='margin-bottom:4px;'>{ln.lstrip('•-* ').strip()}</li>" for ln in note_lines)
            note_html = (
                "<div style='margin:14px 24px 4px 24px;font-style:italic;font-size:12.5px;'>*Note :</div>"
                f"<ul style='margin:2px 24px 12px 44px;font-size:12.5px;'>{li_items}</ul>"
            )

        body_content = f"""
            <div style="text-align:center;font-weight:700;font-size:13.8px;line-height:1.4;margin:14px 24px 14px 24px;">
                <div>{b1}</div>
                <div>{b2}</div>
                <div>{b3}</div>
            </div>
            <p class="lpj-p-indent">
                Dana untuk kegiatan ini berasal dari {sumber}. Dengan rincian dana sebagai berikut :
            </p>
            <table class="dana-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="width:6%;">NO</th>
                        <th rowspan="2" style="width:31%;">KEBUTUHAN KEGIATAN</th>
                        <th rowspan="2" style="width:9%;">JUMLAH</th>
                        <th colspan="2" style="width:32%;">DANA</th>
                        <th rowspan="2" style="width:22%;">KETERANGAN</th>
                    </tr>
                    <tr>
                        <th style="width:16%;">HARGA SATUAN</th>
                        <th style="width:16%;">HARGA TOTAL</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                    <tr>
                        <td colspan="4" style="text-align:center;font-weight:700;background:#F8FAFC;">TOTAL</td>
                        <td style="text-align:right;font-weight:700;white-space:nowrap;background:#F8FAFC;">{fmt_rp_cell(total_dana, is_dash_if_zero=False)}</td>
                        <td style="background:#F8FAFC;"></td>
                    </tr>
                </tbody>
            </table>
            {note_html}
        """
    else:
        sig_bytes = custom_sig_bytes if custom_sig_bytes else assets["sig"]
        sig_b64 = base64.b64encode(sig_bytes).decode("ascii") if sig_bytes else ""

        tgl_laporan_indo = format_hari_tanggal_indo(payload.get("tgl_laporan", date.today()))
        tgl_acara_indo = (payload.get("tgl_acara_manual") or "").strip() or format_hari_tanggal_indo(payload.get("tgl_acara", date.today()))
        waktu_acara = (payload.get("waktu_acara") or "08.30 - selesai").strip()
        tempat_acara = (payload.get("tempat_acara") or "-").strip()
        teks_pembuka = build_paragraf_pembuka(
            payload.get("nama_kegiatan", ""),
            payload.get("tema_kegiatan", ""),
            payload.get("konteks_hut", ""),
        )

        rangkaian_raw = (payload.get("rangkaian_kegiatan") or "").strip()
        rangkaian_items = [ln.strip() for ln in rangkaian_raw.split(NL) if ln.strip()]
        ol_html = ""
        for item_r in rangkaian_items:
            clean_r = item_r
            if "." in clean_r[:4]:
                l_num, r_rest = clean_r.split(".", 1)
                if l_num.strip().isdigit():
                    clean_r = r_rest.strip()
            ol_html += f"<li style='margin-bottom:3px;'>{clean_r}</li>"

        df_items = normalize_items_df(payload.get("items_dana", []), payload.get("ket_tabel_dana", ""))
        n_items = len(df_items)
        merge_ket = bool(payload.get("merge_ket_vertikal", True))
        ket_global = (payload.get("ket_tabel_dana") or "Kas DWP BC Malang & Donatur Anggota").strip()
        total_dana = int(df_items["HARGA TOTAL"].sum()) if not df_items.empty else 0

        rows_html = ""
        for idx_df, r in df_items.iterrows():
            no_str = str(int(r["NO"]))
            keb_str = str(r["KEBUTUHAN KEGIATAN"])
            jml_str = str(int(r["JUMLAH"])) if int(r["JUMLAH"]) > 0 else "-"
            hsat_str = fmt_rp_cell(r["HARGA SATUAN"], is_dash_if_zero=True)
            htot_str = fmt_rp_cell(r["HARGA TOTAL"], is_dash_if_zero=True)
            ket_item = str(r.get("KETERANGAN", ket_global) or ket_global)

            if merge_ket:
                if idx_df == 0:
                    ket_td = f"<td rowspan='{n_items}' style='text-align:center;vertical-align:middle;padding:6px;'>{ket_global}</td>"
                else:
                    ket_td = ""
            else:
                ket_td = f"<td style='text-align:center;vertical-align:middle;padding:4px 6px;'>{ket_item}</td>"

            rows_html += (
                "<tr>"
                f"<td style='text-align:center;'>{no_str}</td>"
                f"<td style='text-align:left;'>{keb_str}</td>"
                f"<td style='text-align:center;'>{jml_str}</td>"
                f"<td style='text-align:right;white-space:nowrap;'>{hsat_str}</td>"
                f"<td style='text-align:right;white-space:nowrap;'>{htot_str}</td>"
                f"{ket_td}"
                "</tr>"
            )

        ucapan_html = ""
        if payload.get("tampilkan_ucapan_donatur", True):
            teks_don = (payload.get("teks_ucapan_donatur") or DEFAULT_UCAPAN_DONATUR).strip()
            if teks_don:
                ucapan_html = f"<p class='lpj-p-indent'>{teks_don}</p>"

        nama_keg = (payload.get("nama_kegiatan") or "Kegiatan Dharma Wanita Persatuan").strip()
        penutup_html = (
            f"<p class='lpj-p-indent'>Demikian laporan pertanggung jawaban {nama_keg} ini kami sampaikan, "
            "atas perhatian dan bantuan yang telah diberikan kami ucapkan terima kasih.</p>"
        )

        catatan_ig = (payload.get("catatan_ig") or "").strip()
        ig_html = ""
        if catatan_ig:
            ig_formatted = catatan_ig.replace("@dwpbcmalang", "<b>@dwpbcmalang</b>")
            ig_html = f"<p style='margin:10px 0 12px 24px;font-size:12.5px;font-style:italic;'>{ig_formatted}</p>"

        jab_1 = (payload.get("jabatan_baris_1") or "Ketua Dharma Wanita Persatuan").strip()
        jab_2 = (payload.get("jabatan_baris_2") or "KPPBC Tipe Madya Cukai Malang").strip()
        nama_ttd = (payload.get("nama_penandatangan") or "Ny. Rini Johan Pandores").strip()

        if payload.get("tampilkan_ttd", True) and sig_b64:
            sig_img_html = f"<img src='data:image/jpeg;base64,{sig_b64}' style='height:58px;margin:4px 0;display:block;'/>"
        else:
            sig_img_html = "<div style='height:54px;'></div>"

        body_content = f"""
            <div style="text-align:right;margin:8px 28px 14px 24px;">{tgl_laporan_indo}</div>
            <div style="margin:6px 24px 10px 24px;font-style:italic;">Assalamu’alaikum Wr. Wb.</div>
            <p class="lpj-p-indent">{teks_pembuka}</p>
            <table class="info-table">
                <tr><td style="width:110px;">Hari, Tanggal</td><td>: {tgl_acara_indo}</td></tr>
                <tr><td>Waktu</td><td>: {waktu_acara}</td></tr>
                <tr><td>Tempat</td><td>: {tempat_acara}</td></tr>
            </table>
            <div style="margin:8px 24px 4px 64px;">Dengan rangkaian kegiatan sebagai berikut :</div>
            <ol style="margin:2px 28px 10px 82px;padding-left:14px;">{ol_html}</ol>
            <p class="lpj-p-indent">
                Dana untuk kegiatan ini berasal dari {payload.get("sumber_dana", "")}. Dengan rincian dana sebagai berikut :
            </p>
            <table class="dana-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="width:6%;">NO</th>
                        <th rowspan="2" style="width:31%;">KEBUTUHAN KEGIATAN</th>
                        <th rowspan="2" style="width:9%;">JUMLAH</th>
                        <th colspan="2" style="width:32%;">DANA</th>
                        <th rowspan="2" style="width:22%;">KETERANGAN</th>
                    </tr>
                    <tr>
                        <th style="width:16%;">HARGA SATUAN</th>
                        <th style="width:16%;">HARGA TOTAL</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                    <tr>
                        <td colspan="4" style="text-align:center;font-weight:700;background:#F8FAFC;">TOTAL</td>
                        <td style="text-align:right;font-weight:700;white-space:nowrap;background:#F8FAFC;">{fmt_rp_cell(total_dana, is_dash_if_zero=False)}</td>
                        <td style="background:#F8FAFC;"></td>
                    </tr>
                </tbody>
            </table>
            {ucapan_html}
            {penutup_html}
            <div style="margin:8px 24px 6px 24px;font-style:italic;">Wassalamu'alaikum Wr. Wb.</div>
            {ig_html}
            <div style="margin-left:58%;margin-top:14px;line-height:1.35;">
                <div>{jab_1}</div>
                <div>{jab_2}</div>
                {sig_img_html}
                <div style="font-weight:600;">{nama_ttd}</div>
            </div>
        """

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8"/>
    <style>
        :root {{ color-scheme: light !important; }}
        body {{
            background: #E2E8F0;
            margin: 0;
            padding: 16px 8px;
            font-family: Arial, Helvetica, sans-serif;
            color: #000000 !important;
        }}
        .toolbar-print {{
            max-width: 215mm;
            margin: 0 auto 12px auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #FFFFFF;
            padding: 10px 16px;
            border-radius: 10px;
            border: 1px solid #CBD5E1;
            box-shadow: 0 2px 6px rgba(15,23,42,0.06);
        }}
        .btn-print {{
            background: linear-gradient(135deg, #2B98F0, #1D7FE0);
            color: #FFFFFF !important;
            border: none;
            padding: 9px 16px;
            border-radius: 8px;
            font-weight: 700;
            font-size: 13.5px;
            cursor: pointer;
        }}
        .a4-sheet {{
            max-width: 215mm;
            min-height: 279mm;
            margin: 0 auto;
            background: #FFFFFF !important;
            color: #000000 !important;
            padding: 12mm 14mm 15mm 16mm;
            box-shadow: 0 8px 28px rgba(15,23,42,0.15);
            box-sizing: border-box;
            font-size: 13.5px;
            line-height: 1.42;
        }}
        .kop-table {{
            width: 100%;
            border-collapse: collapse;
            border-bottom: 2.2px solid #000000;
            margin-bottom: 12px;
        }}
        .kop-table td {{
            border: none !important;
            vertical-align: middle;
            padding-bottom: 6px;
        }}
        .lpj-p-indent {{
            text-align: justify;
            text-indent: 36px;
            margin: 6px 24px;
        }}
        .info-table {{
            margin: 6px 0 6px 64px;
            border-collapse: collapse;
        }}
        .info-table td {{
            padding: 1.5px 6px 1.5px 0;
            vertical-align: top;
            font-size: 13.5px;
        }}
        .dana-table {{
            width: calc(100% - 24px);
            margin: 8px 12px 12px 12px;
            border-collapse: collapse;
            font-size: 12.8px;
        }}
        .dana-table th, .dana-table td {{
            border: 1px solid #475569;
            padding: 4px 6px;
            color: #000000 !important;
        }}
        .dana-table th {{
            background: #F8FAFC;
            font-weight: 700;
            text-align: center;
        }}
        @media print {{
            body {{ background: #FFFFFF; padding: 0; }}
            .toolbar-print {{ display: none !important; }}
            .a4-sheet {{ box-shadow: none; margin: 0; max-width: 100%; padding: 8mm 10mm; }}
        }}
    </style>
    </head>
    <body>
        <div class="toolbar-print">
            <div style="font-size:13px;font-weight:700;color:#1E293B;">
                📄 Pratinjau Kertas A4 Resmi — Mode Dokumen: {doc_mode} DWP KPPBC TMC Malang
            </div>
            <button class="btn-print" onclick="window.print()">🖨️ Cetak / Simpan PDF via Browser</button>
        </div>
        <div class="a4-sheet">
            {kop_html}
            {body_content}
        </div>
    </body>
    </html>
    """


# ==============================================================================
# BAGIAN 6: SPESIFIKASI CSS "FILY" MODERN ADMIN UI (ANTI DARK-MODE & MOBILE-FIRST)
# ==============================================================================

def inject_fily_admin_css():
    st.markdown(
        """
        <style>
        :root {
            color-scheme: light !important;
        }
        html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
            background-color: #F1F5F9 !important;
            color: #0F172A !important;
            font-family: 'Inter', 'Segoe UI', Roboto, sans-serif;
        }
        [data-testid="stHeader"] {
            background-color: rgba(241, 245, 249, 0.90) !important;
        }
        .block-container {
            padding-top: 1.4rem !important;
            padding-bottom: 2.5rem !important;
            max-width: 1420px !important;
        }
        /* Dark Slate Navy Sidebar (#1E2235) */
        [data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {
            background-color: #1E2235 !important;
            border-right: 1px solid #2D334D !important;
        }
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] h4,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] caption {
            color: #F8FAFC !important;
            -webkit-text-fill-color: #F8FAFC !important;
        }
        /* Navigasi Radio di Sidebar tampil seperti Menu Web Modern */
        [data-testid="stSidebar"] div[role="radiogroup"] > label {
            background-color: rgba(255, 255, 255, 0.04) !important;
            border: 1px solid rgba(255, 255, 255, 0.08) !important;
            border-left: 4px solid transparent !important;
            border-radius: 10px !important;
            padding: 10px 14px !important;
            margin-bottom: 6px !important;
            min-height: 46px !important;
            transition: all 0.18s ease !important;
            cursor: pointer !important;
        }
        [data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
            background-color: #282E47 !important;
            border-left: 4px solid #2B98F0 !important;
        }
        [data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"],
        [data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
            background: linear-gradient(90deg, rgba(43,152,240,0.25) 0%, rgba(40,46,71,0.95) 100%) !important;
            border: 1px solid rgba(43, 152, 240, 0.45) !important;
            border-left: 4px solid #00D09C !important;
            font-weight: 700 !important;
        }
        /* Input di dalam Expander Sidebar tetap berlatar Putih & Teks Gelap */
        [data-testid="stSidebar"] input,
        [data-testid="stSidebar"] textarea,
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
        }
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
        }
        /* Area Utama: Paksa semua teks gelap pekat #0F172A */
        [data-testid="stMain"] p,
        [data-testid="stMain"] span,
        [data-testid="stMain"] label,
        [data-testid="stMain"] h1,
        [data-testid="stMain"] h2,
        [data-testid="stMain"] h3,
        [data-testid="stMain"] h4,
        [data-testid="stMain"] li,
        [data-testid="stMain"] td,
        [data-testid="stMain"] th {
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
        }
        /* Input Ergonomis Mobile-First (Anti Auto-Zoom HP: 16px, min-height 46px) */
        [data-testid="stMain"] input,
        [data-testid="stMain"] textarea,
        [data-testid="stMain"] select,
        [data-testid="stMain"] div[data-baseweb="input"] input,
        [data-testid="stMain"] div[data-baseweb="textarea"] textarea,
        [data-testid="stMain"] div[data-baseweb="select"] > div {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
            font-size: 16px !important;
            min-height: 46px !important;
            border-radius: 8px !important;
            border: 1.5px solid #CBD5E1 !important;
        }
        div[data-baseweb="popover"], ul[data-baseweb="menu"], li[role="option"] {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
            font-size: 16px !important;
        }
        /* Tombol Utama Modern */
        [data-testid="stMain"] .stButton > button {
            min-height: 46px !important;
            border-radius: 10px !important;
            font-size: 14.5px !important;
            font-weight: 700 !important;
            background: linear-gradient(135deg, #2B98F0 0%, #1D7FE0 100%) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(43, 152, 240, 0.22) !important;
            transition: all 0.18s ease !important;
        }
        [data-testid="stMain"] .stButton > button p,
        [data-testid="stMain"] .stButton > button span {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        [data-testid="stMain"] .stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 16px rgba(43, 152, 240, 0.32) !important;
        }
        [data-testid="stSidebar"] .stButton > button {
            min-height: 42px !important;
            border-radius: 9px !important;
            font-size: 13.5px !important;
            font-weight: 700 !important;
            background: linear-gradient(135deg, #2B98F0 0%, #1D7FE0 100%) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border: 1px solid rgba(255,255,255,0.18) !important;
        }
        /* Top Welcome Bar & 4 Pill KPI Cards ("Fily" Style) */
        .fily-topbar {
            background: #FFFFFF !important;
            padding: 14px 22px;
            border-radius: 12px;
            border: 1px solid #E2E8F0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
            margin-bottom: 18px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        }
        .fily-pill {
            border-radius: 999px;
            padding: 15px 22px;
            color: #FFFFFF !important;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.10);
            margin-bottom: 16px;
        }
        .fily-pill * {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        .fily-pill-icon {
            width: 48px;
            height: 48px;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.22);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
            flex-shrink: 0;
        }
        .pill-emerald { background: linear-gradient(135deg, #00D09C 0%, #00B386 100%); }
        .pill-blue    { background: linear-gradient(135deg, #2B98F0 0%, #1D7FE0 100%); }
        .pill-rose    { background: linear-gradient(135deg, #F85370 0%, #E63956 100%); }
        .pill-amber   { background: linear-gradient(135deg, #FBA01D 0%, #F58800 100%); }

        .fily-card {
            background: #FFFFFF !important;
            border-radius: 12px;
            border: 1px solid #E2E8F0;
            padding: 20px 22px;
            margin-bottom: 16px;
            box-shadow: 0 2px 10px rgba(15, 23, 42, 0.03);
        }
        .fily-section-hdr {
            background: #FFFFFF !important;
            border-left: 5px solid #2B98F0;
            border-radius: 10px;
            padding: 12px 18px;
            margin-bottom: 16px;
            border-top: 1px solid #E2E8F0;
            border-right: 1px solid #E2E8F0;
            border-bottom: 1px solid #E2E8F0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# BAGIAN 7: MANAJEMEN STATE, PRESET LOADER & 1-CLICK RAB-TO-LPJ CONVERTER
# ==============================================================================

def load_rab_into_state(data: dict):
    """Memuat seluruh field RAB ke st.session_state."""
    st.session_state["id_rab"] = data.get("id_rab", "RAB-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
    st.session_state["baris_1_judul"] = data.get("baris_1_judul", "RANCANGAN ANGGARAN BIAYA")
    st.session_state["nama_kegiatan_rab"] = data.get("nama_kegiatan_rab", "")
    st.session_state["lokasi_acara_rab"] = data.get("lokasi_acara_rab", "")
    st.session_state["sumber_dana_rab"] = data.get("sumber_dana_rab", "Kas Dharma Wanita KPPBC TMC Malang")
    st.session_state["ket_tabel_rab"] = data.get("ket_tabel_rab", "")
    st.session_state["merge_ket_rab"] = bool(data.get("merge_ket_rab", False))
    st.session_state["catatan_note_rab"] = data.get("catatan_note_rab", "")
    st.session_state["status_konversi"] = data.get("status_konversi", "BELUM DIKONVERSI")
    st.session_state["df_items_rab"] = normalize_items_df(data.get("items_rab", []), st.session_state["ket_tabel_rab"])


def load_lpj_into_state(data: dict):
    """Memuat seluruh field LPJ ke st.session_state."""
    st.session_state["id_lpj"] = data.get("id_lpj", "LPJ-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
    st.session_state["id_rab_ref"] = data.get("id_rab_ref", "")

    tgl_lap = data.get("tgl_laporan", date(2026, 8, 13))
    if isinstance(tgl_lap, str):
        try:
            tgl_lap = datetime.strptime(tgl_lap[:10], "%Y-%m-%d").date()
        except Exception:
            tgl_lap = date(2026, 8, 13)
    st.session_state["tgl_laporan"] = tgl_lap

    st.session_state["nama_kegiatan"] = data.get("nama_kegiatan", "")
    st.session_state["tema_kegiatan"] = data.get("tema_kegiatan", "")
    st.session_state["konteks_hut"] = data.get("konteks_hut", "")

    tgl_acr = data.get("tgl_acara", date(2026, 8, 10))
    if isinstance(tgl_acr, str):
        try:
            tgl_acr = datetime.strptime(tgl_acr[:10], "%Y-%m-%d").date()
        except Exception:
            tgl_acr = date(2026, 8, 10)
    st.session_state["tgl_acara"] = tgl_acr

    st.session_state["tgl_acara_manual"] = data.get("tgl_acara_manual", "")
    st.session_state["waktu_acara"] = data.get("waktu_acara", "08.30 - selesai")
    st.session_state["tempat_acara"] = data.get("tempat_acara", "")
    st.session_state["rangkaian_kegiatan"] = data.get("rangkaian_kegiatan", "")
    st.session_state["sumber_dana"] = data.get("sumber_dana", "Kas Dharma Wanita KPPBC TMC Malang dan donatur")
    st.session_state["ket_tabel_dana"] = data.get("ket_tabel_dana", "Kas DWP BC Malang & Donatur Anggota")
    st.session_state["merge_ket_vertikal"] = bool(data.get("merge_ket_vertikal", True))
    st.session_state["tampilkan_ucapan_donatur"] = bool(data.get("tampilkan_ucapan_donatur", True))
    st.session_state["teks_ucapan_donatur"] = data.get("teks_ucapan_donatur", DEFAULT_UCAPAN_DONATUR)
    st.session_state["catatan_ig"] = data.get("catatan_ig", "*Note : dokumentasi kegiatan dapat dilihat di Instagram @dwpbcmalang")
    st.session_state["jabatan_baris_1"] = data.get("jabatan_baris_1", "Ketua Dharma Wanita Persatuan")
    st.session_state["jabatan_baris_2"] = data.get("jabatan_baris_2", "KPPBC Tipe Madya Cukai Malang")
    st.session_state["nama_penandatangan"] = data.get("nama_penandatangan", "Ny. Rini Johan Pandores")
    st.session_state["tampilkan_ttd"] = bool(data.get("tampilkan_ttd", True))
    st.session_state["df_items_dana"] = normalize_items_df(data.get("items_dana", []), st.session_state["ket_tabel_dana"])


def convert_active_rab_to_lpj(df_rab_current: pd.DataFrame = None):
    """
    1-Klik Konversi RAB ke LPJ:
    Menarik Nama Kegiatan, Lokasi Acara, Sumber Dana, dan seluruh baris anggaran RAB
    langsung menjadi draf LPJ siap pakai tanpa ketik ulang!
    """
    raw_nama_rab = (st.session_state.get("nama_kegiatan_rab") or "").strip()
    if raw_nama_rab.upper().startswith("PELAKSANAAN "):
        raw_nama_rab = raw_nama_rab[12:].strip()

    # Ubah kapitalisasi judul RAB menjadi kalimat LPJ yang elegan
    nama_lpj = raw_nama_rab.title()
    for acr in ["DWP", "DJBC", "KPPBC", "TMC", "TMP", "HUT", "RI", "II", "III"]:
        nama_lpj = nama_lpj.replace(acr.title(), acr)
    nama_lpj = nama_lpj.replace("Di Lingkungan", "di lingkungan")

    lok_rab = (st.session_state.get("lokasi_acara_rab") or "").strip().strip("()")
    if lok_rab.upper().startswith("DI "):
        lok_rab = lok_rab[3:].strip()
    tempat_lpj = lok_rab.title()
    for acr in ["KPPBC", "TMC", "TMP", "DJBC", "C", "A", "B"]:
        tempat_lpj = tempat_lpj.replace(f" {acr.title()} ", f" {acr} ").replace(f"{acr.title()} ", f"{acr} ")

    sumber_rab = (st.session_state.get("sumber_dana_rab") or "Kas Dharma Wanita KPPBC TMC Malang").strip()
    if "donatur" not in sumber_rab.lower():
        sumber_lpj = sumber_rab + " dan donatur"
    else:
        sumber_lpj = sumber_rab

    ket_lpj = st.session_state.get("ket_tabel_dana") or "Kas DWP BC Malang & Donatur Anggota"
    df_src = df_rab_current if df_rab_current is not None else st.session_state.get("df_items_rab", pd.DataFrame())
    df_lpj_new = normalize_items_df(df_src, ket_lpj)
    df_lpj_new["KETERANGAN"] = ket_lpj

    st.session_state["id_lpj"] = "LPJ-" + datetime.now().strftime("%Y%m%d-%H%M%S")
    st.session_state["id_rab_ref"] = st.session_state.get("id_rab", "RAB-AKTIF")
    st.session_state["nama_kegiatan"] = nama_lpj
    st.session_state["tempat_acara"] = tempat_lpj
    st.session_state["sumber_dana"] = sumber_lpj
    st.session_state["df_items_dana"] = df_lpj_new
    st.session_state["status_konversi"] = f"SUDAH DIKONVERSI KE {st.session_state['id_lpj']}"
    st.session_state["preview_doc_mode"] = "LPJ"


def init_app_state():
    if "app_v5_initialized" not in st.session_state:
        load_rab_into_state(PRESET_RAB_JEMBER)
        load_lpj_into_state(PRESET_LPJ_BLITAR_2026)
        st.session_state["local_db_rab"] = []
        st.session_state["local_db_lpj"] = []
        st.session_state["local_rekap_dana"] = []
        st.session_state["active_menu"] = MENU_DASHBOARD
        st.session_state["preview_doc_mode"] = "LPJ"
        st.session_state["app_v5_initialized"] = True


def collect_rab_payload(df_rab_override: pd.DataFrame = None) -> dict:
    df_use = df_rab_override if df_rab_override is not None else st.session_state.get("df_items_rab", pd.DataFrame())
    return {
        "id_rab": st.session_state.get("id_rab", "RAB-2026-JEMBER"),
        "baris_1_judul": st.session_state.get("baris_1_judul", "RANCANGAN ANGGARAN BIAYA"),
        "nama_kegiatan_rab": st.session_state.get("nama_kegiatan_rab", ""),
        "lokasi_acara_rab": st.session_state.get("lokasi_acara_rab", ""),
        "sumber_dana_rab": st.session_state.get("sumber_dana_rab", ""),
        "ket_tabel_rab": st.session_state.get("ket_tabel_rab", ""),
        "merge_ket_rab": st.session_state.get("merge_ket_rab", False),
        "catatan_note_rab": st.session_state.get("catatan_note_rab", ""),
        "status_konversi": st.session_state.get("status_konversi", "BELUM DIKONVERSI"),
        "items_rab": df_use.to_dict(orient="records"),
    }


def collect_lpj_payload(df_lpj_override: pd.DataFrame = None) -> dict:
    df_use = df_lpj_override if df_lpj_override is not None else st.session_state.get("df_items_dana", pd.DataFrame())
    return {
        "id_lpj": st.session_state.get("id_lpj", "LPJ-20260813-BLITAR"),
        "id_rab_ref": st.session_state.get("id_rab_ref", ""),
        "tgl_laporan": st.session_state.get("tgl_laporan", date(2026, 8, 13)),
        "nama_kegiatan": st.session_state.get("nama_kegiatan", ""),
        "tema_kegiatan": st.session_state.get("tema_kegiatan", ""),
        "konteks_hut": st.session_state.get("konteks_hut", ""),
        "tgl_acara": st.session_state.get("tgl_acara", date(2026, 8, 10)),
        "tgl_acara_manual": st.session_state.get("tgl_acara_manual", ""),
        "waktu_acara": st.session_state.get("waktu_acara", "08.30 - selesai"),
        "tempat_acara": st.session_state.get("tempat_acara", ""),
        "rangkaian_kegiatan": st.session_state.get("rangkaian_kegiatan", ""),
        "sumber_dana": st.session_state.get("sumber_dana", ""),
        "ket_tabel_dana": st.session_state.get("ket_tabel_dana", ""),
        "merge_ket_vertikal": st.session_state.get("merge_ket_vertikal", True),
        "tampilkan_ucapan_donatur": st.session_state.get("tampilkan_ucapan_donatur", True),
        "teks_ucapan_donatur": st.session_state.get("teks_ucapan_donatur", DEFAULT_UCAPAN_DONATUR),
        "catatan_ig": st.session_state.get("catatan_ig", ""),
        "jabatan_baris_1": st.session_state.get("jabatan_baris_1", "Ketua Dharma Wanita Persatuan"),
        "jabatan_baris_2": st.session_state.get("jabatan_baris_2", "KPPBC Tipe Madya Cukai Malang"),
        "nama_penandatangan": st.session_state.get("nama_penandatangan", "Ny. Rini Johan Pandores"),
        "tampilkan_ttd": st.session_state.get("tampilkan_ttd", True),
        "items_dana": df_use.to_dict(orient="records"),
    }


def switch_menu(target_menu: str, preview_mode: str = None):
    st.session_state[" _nav_target"] = target_menu
    if preview_mode:
        st.session_state["preview_doc_mode"] = preview_mode
    st.rerun()


# ==============================================================================
# BAGIAN 8: MAIN APPLICATION (SINGLE ACTIVE SIDEBAR VIEW — ANTI-LAG)
# ==============================================================================

def main():
    inject_fily_admin_css()
    init_app_state()

    if " _nav_target" in st.session_state:
        st.session_state["active_menu"] = st.session_state.pop(" _nav_target")

    # --- SIDEBAR: DARK SLATE NAVY ("FILY" ADMIN NAVIGATION) ---
    with st.sidebar:
        st.markdown(
            """
            <div style="display:flex;align-items:center;gap:12px;padding:6px 4px 16px 4px;border-bottom:1px solid #2D334D;margin-bottom:14px;">
                <div style="width:42px;height:42px;border-radius:10px;background:linear-gradient(135deg,#2B98F0,#00D09C);display:flex;align-items:center;justify-content:center;font-size:22px;box-shadow:0 4px 12px rgba(43,152,240,0.3);">
                    🏛️
                </div>
                <div>
                    <div style="font-weight:800;font-size:16px;letter-spacing:0.4px;color:#FFFFFF;">DWP BC MALANG</div>
                    <div style="font-size:11.5px;color:#94A3B8;">Enterprise RAB & LPJ Suite V5.0</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.caption("MENU NAVIGASI UTAMA")
        active_menu = st.radio(
            "Navigasi Utama",
            options=ALL_MENUS,
            key="active_menu",
            label_visibility="collapsed",
        )

        st.markdown('<div style="border-top:1px solid #2D334D;margin:14px 0 10px 0;"></div>', unsafe_allow_html=True)
        st.caption("PINTASAN PRESET RESMI 1-KLIK")

        if st.button("📋 Muat Preset RAB Jember", use_container_width=True):
            load_rab_into_state(PRESET_RAB_JEMBER)
            st.toast("Preset RAB DWP Gabungan di Jember berhasil dimuat!", icon="📋")
            switch_menu(MENU_RAB)

        if st.button("📝 Muat Preset LPJ Blitar 2026", use_container_width=True):
            load_lpj_into_state(PRESET_LPJ_BLITAR_2026)
            st.toast("Preset LPJ DWP Gabungan di Blitar (2026) berhasil dimuat!", icon="📝")
            switch_menu(MENU_LPJ)

        st.markdown('<div style="border-top:1px solid #2D334D;margin:14px 0 10px 0;"></div>', unsafe_allow_html=True)
        with st.expander("⚙️ KONFIGURASI SISTEM & CLOUD", expanded=False):
            sheet_id_in = st.text_input("ID Google Sheets", value=DEFAULT_SPREADSHEET_ID, key="cfg_sheet_id")
            up_sa = st.file_uploader("Upload JSON Key Cloud", type=["json"], key="up_sa_json")
            if up_sa is not None:
                st.session_state["creds_json_str"] = up_sa.getvalue().decode("utf-8", errors="ignore")
            creds_str = st.session_state.get("creds_json_str", "")

            up_tpl = st.file_uploader("Upload Template .docx (Opsional)", type=["docx"], key="up_tpl_sidebar")
            if up_tpl is not None:
                st.session_state["uploaded_template_bytes"] = up_tpl.read()
                st.success("Template .docx aktif!")

            if st.button("🔄 Sinkronkan Cloud", use_container_width=True):
                fetch_sheet_bundle_cached.clear()
                st.rerun()

        sheet_id_active = st.session_state.get("cfg_sheet_id", DEFAULT_SPREADSHEET_ID)
        creds_active = st.session_state.get("creds_json_str", "")
        bundle = get_sheet_bundle(sheet_id_active, creds_active)

        status_dot = "🟢 Cloud Terhubung" if bundle.get("ok") else "🟡 Mode Lokal Aktif"
        st.caption(f"Status: **{status_dot}**")

    # Hitung Metrik Utama Secara Ringan (Tanpa Memanggil LibreOffice di Luar Halaman Preview)
    df_rab_now = normalize_items_df(st.session_state.get("df_items_rab", pd.DataFrame()), st.session_state.get("ket_tabel_rab", ""))
    df_lpj_now = normalize_items_df(st.session_state.get("df_items_dana", pd.DataFrame()), st.session_state.get("ket_tabel_dana", ""))

    tot_rab_val = int(df_rab_now["HARGA TOTAL"].sum()) if not df_rab_now.empty else 0
    tot_lpj_val = int(df_lpj_now["HARGA TOTAL"].sum()) if not df_lpj_now.empty else 0
    selisih_val = tot_lpj_val - tot_rab_val
    arsip_count = len(bundle.get("rab_rows", [])) + len(bundle.get("lpj_rows", []))

    # --- TOP WELCOME BAR ---
    tb_col1, tb_col2 = st.columns([2.6, 1.1])
    with tb_col1:
        st.markdown(
            f"""
            <div class="fily-topbar">
                <div>
                    <div style="font-size:18px;font-weight:800;color:#0F172A;">
                        👋 Hi, Selamat Datang Pengurus DWP KPPBC TMC Malang!
                    </div>
                    <div style="font-size:13px;color:#475569;margin-top:2px;">
                        Alur Terpadu: <b>1. Susun RAB</b> → <b>2. Konversi 1-Klik ke LPJ</b> → <b>3. Cetak Word & PDF Presisi 100%</b>
                    </div>
                </div>
                <div style="font-size:12.5px;font-weight:700;background:#EFF6FF;color:#1D4ED8;padding:6px 12px;border-radius:8px;border:1px solid #BFDBFE;">
                    {status_dot}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with tb_col2:
        if st.button("🚀 Konversi RAB Aktif → LPJ (1-Klik)", use_container_width=True):
            convert_active_rab_to_lpj(df_rab_now)
            st.toast("Seluruh data RAB berhasil dikonversi menjadi draf LPJ!", icon="🚀")
            switch_menu(MENU_LPJ)

    # ==========================================================================
    # VIEW 0: 🏠 DASHBOARD UTAMA ("FILY" 4 PILL KPI CARDS & TABEL KOMPARASI)
    # ==========================================================================
    if active_menu == MENU_DASHBOARD:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                f"""
                <div class="fily-pill pill-emerald">
                    <div class="fily-pill-icon">↗</div>
                    <div style="text-align:right;">
                        <div style="font-size:12px;font-weight:700;opacity:0.92;">TOTAL RENCANA (RAB)</div>
                        <div style="font-size:20px;font-weight:800;">{fmt_rp_cell(tot_rab_val, False)}</div>
                        <div style="font-size:11.5px;opacity:0.9;">{len(df_rab_now)} Item Anggaran</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f"""
                <div class="fily-pill pill-blue">
                    <div class="fily-pill-icon">🛍️</div>
                    <div style="text-align:right;">
                        <div style="font-size:12px;font-weight:700;opacity:0.92;">TOTAL REALISASI (LPJ)</div>
                        <div style="font-size:20px;font-weight:800;">{fmt_rp_cell(tot_lpj_val, False)}</div>
                        <div style="font-size:11.5px;opacity:0.9;">{len(df_lpj_now)} Item Realisasi</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with k3:
            selisih_label = ("+" if selisih_val > 0 else "") + fmt_ribuan(selisih_val)
            st.markdown(
                f"""
                <div class="fily-pill pill-rose">
                    <div class="fily-pill-icon">🎯</div>
                    <div style="text-align:right;">
                        <div style="font-size:12px;font-weight:700;opacity:0.92;">SELISIH (LPJ vs RAB)</div>
                        <div style="font-size:20px;font-weight:800;">Rp {selisih_label}</div>
                        <div style="font-size:11.5px;opacity:0.9;">Analisis Realisasi Dana</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f"""
                <div class="fily-pill pill-amber">
                    <div class="fily-pill-icon">📊</div>
                    <div style="text-align:right;">
                        <div style="font-size:12px;font-weight:700;opacity:0.92;">ARSIP DOKUMEN</div>
                        <div style="font-size:20px;font-weight:800;">{arsip_count} Dokumen</div>
                        <div style="font-size:11.5px;opacity:0.9;">Tersimpan di Database</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        col_comp, col_short = st.columns([1.85, 1.15])
        with col_comp:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### 📊 Tabel Komparasi Anggaran (Rencana RAB vs Realisasi LPJ)")
            max_len = max(len(df_rab_now), len(df_lpj_now))
            comp_rows = []
            for i in range(max_len):
                r_rab = df_rab_now.iloc[i] if i < len(df_rab_now) else None
                r_lpj = df_lpj_now.iloc[i] if i < len(df_lpj_now) else None
                comp_rows.append({
                    "NO": i + 1,
                    "ITEM RENCANA (RAB)": str(r_rab["KEBUTUHAN KEGIATAN"]) if r_rab is not None else "-",
                    "ANGGARAN RAB": fmt_rp_cell(r_rab["HARGA TOTAL"], True) if r_rab is not None else "-",
                    "ITEM REALISASI (LPJ)": str(r_lpj["KEBUTUHAN KEGIATAN"]) if r_lpj is not None else "-",
                    "REALISASI LPJ": fmt_rp_cell(r_lpj["HARGA TOTAL"], True) if r_lpj is not None else "-",
                })
            st.dataframe(pd.DataFrame(comp_rows), use_container_width=True, hide_index=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with col_short:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### ⚡ Ringkasan Dokumen Aktif & Pintasan Kerja")
            st.markdown(
                f"""
                - **RAB Aktif (`{st.session_state.get('id_rab')}`)**:
                  {st.session_state.get('lokasi_acara_rab')} — **{fmt_rp_cell(tot_rab_val, False)}**
                - **LPJ Aktif (`{st.session_state.get('id_lpj')}`)**:
                  {st.session_state.get('tempat_acara')} — **{fmt_rp_cell(tot_lpj_val, False)}**
                - **Status Konversi RAB → LPJ**:
                  `{st.session_state.get('status_konversi')}`
                """
            )
            if st.button("📋 Buka Editor Penyusunan RAB", use_container_width=True):
                switch_menu(MENU_RAB)
            if st.button("📝 Buka Editor Pembuatan LPJ", use_container_width=True):
                switch_menu(MENU_LPJ)
            if st.button("🖨️ Buka Preview Kertas A4 & Download Word/PDF", use_container_width=True):
                switch_menu(MENU_PREVIEW)
            st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================================================
    # VIEW 1: 📋 1. SUSUN RAB (PRA-KEGIATAN)
    # ==========================================================================
    elif active_menu == MENU_RAB:
        st.markdown(
            """
            <div class="fily-section-hdr">
                <div>
                    <div style="font-size:17px;font-weight:800;color:#0F172A;">📋 MODUL PENYUSUNAN RAB (PRA-KEGIATAN)</div>
                    <div style="font-size:12.5px;color:#475569;">Sesuai Pakem Resmi: RAB DWP GABUNGAN DI JEMBER.docx</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        c_rab1, c_rab2 = st.columns([1.15, 1.15])
        with c_rab1:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### 📌 CARD 1: Identitas & Judul 3 Baris RAB")
            st.text_input("Baris 1 Judul (Tetap)", key="baris_1_judul")
            st.text_area("Baris 2 Nama Kegiatan (Otomatis Kapital di Dokumen)", key="nama_kegiatan_rab", height=76)
            st.text_input("Baris 3 Lokasi Acara (Contoh: (DI KPPBC TMP C JEMBER))", key="lokasi_acara_rab")
            st.text_input("Sumber Dana (Kalimat Pengantar)", key="sumber_dana_rab")
            ck1, ck2 = st.columns([1.4, 1])
            with ck1:
                st.text_input("Keterangan Kolom Tabel RAB (Opsional)", key="ket_tabel_rab")
            with ck2:
                st.write("")
                st.checkbox("Merge Kolom Keterangan", key="merge_ket_rab")
            st.markdown("</div>", unsafe_allow_html=True)

        with c_rab2:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### ⚡ CARD 2: Smart Paste WA & Input Cepat RAB")
            mode_in_rab = st.radio(
                "Pilih Metode Input Cepat:",
                options=["Smart Paste WhatsApp", "Form Tambah 1 Baris"],
                horizontal=True,
                key="rab_input_mode",
            )
            if mode_in_rab == "Smart Paste WhatsApp":
                wa_rab = st.text_area(
                    "Tempel rincian anggaran dari WA (Contoh: `Iuran DWP 7 total 450.000`, `Doorprize 7 x 50.000`, `Mobil 2 -`):",
                    height=115,
                    key="wa_paste_rab",
                )
                b_p1, b_p2 = st.columns(2)
                with b_p1:
                    if st.button("🔮 Tambah ke Tabel RAB", use_container_width=True):
                        items_new = smart_paste_parser(wa_rab, st.session_state.get("ket_tabel_rab", ""))
                        if items_new:
                            cur = st.session_state["df_items_rab"].to_dict(orient="records")
                            cur.extend(items_new)
                            st.session_state["df_items_rab"] = normalize_items_df(cur, st.session_state.get("ket_tabel_rab", ""))
                            st.toast(f"{len(items_new)} item ditambahkan ke RAB!", icon="✅")
                            st.rerun()
                with b_p2:
                    if st.button("🔄 Ganti Seluruh Tabel RAB", use_container_width=True):
                        items_new = smart_paste_parser(wa_rab, st.session_state.get("ket_tabel_rab", ""))
                        if items_new:
                            st.session_state["df_items_rab"] = normalize_items_df(items_new, st.session_state.get("ket_tabel_rab", ""))
                            st.toast("Tabel RAB berhasil diganti!", icon="🔄")
                            st.rerun()
            else:
                qr_nama = st.text_input("Kebutuhan Kegiatan RAB", key="qr_nama")
                q1, q2 = st.columns([0.8, 1.4])
                with q1:
                    qr_jml = st.number_input("Jumlah", min_value=1, value=1, step=1, key="qr_jml")
                with q2:
                    qr_tipe = st.selectbox(
                        "Mode Perhitungan",
                        ["Normal (Qty x Harga Satuan)", "Borongan (Langsung Harga Total)", "Fasilitas / In-Kind (Rp -)"],
                        key="qr_tipe",
                    )
                qr_nom = st.text_input("Nominal (Rp)", value="0", key="qr_nom")
                if st.button("➕ Tambahkan Item ke RAB", use_container_width=True):
                    if qr_nama.strip():
                        v_nom = parse_num(qr_nom)
                        if qr_tipe.startswith("Normal"):
                            hs, ht = v_nom, int(qr_jml) * v_nom
                        elif qr_tipe.startswith("Borongan"):
                            hs, ht = 0, v_nom
                        else:
                            hs, ht = 0, 0
                        cur = st.session_state["df_items_rab"].to_dict(orient="records")
                        cur.append({
                            "KEBUTUHAN KEGIATAN": qr_nama.strip(),
                            "JUMLAH": int(qr_jml),
                            "HARGA SATUAN": hs,
                            "HARGA TOTAL": ht,
                            "KETERANGAN": st.session_state.get("ket_tabel_rab", ""),
                        })
                        st.session_state["df_items_rab"] = normalize_items_df(cur, st.session_state.get("ket_tabel_rab", ""))
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        # CARD 3: Tabel Rincian Anggaran Interaktif (Anti Cursor-Jump)
        st.markdown('<div class="fily-card">', unsafe_allow_html=True)
        hdr_t1, hdr_t2 = st.columns([2.2, 1])
        with hdr_t1:
            st.markdown("#### 💰 CARD 3: Tabel Rincian Anggaran Interaktif (Logika 3-Mode Otomatis)")
        with hdr_t2:
            if st.button("🔄 Hitung Ulang & Rapikan Nomor Urut RAB", use_container_width=True):
                st.session_state["df_items_rab"] = normalize_items_df(
                    st.session_state.get("df_items_rab", pd.DataFrame()),
                    st.session_state.get("ket_tabel_rab", ""),
                )
                st.rerun()

        edited_rab_raw = st.data_editor(
            st.session_state["df_items_rab"],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "NO": st.column_config.NumberColumn("NO", width="small", disabled=True),
                "KEBUTUHAN KEGIATAN": st.column_config.TextColumn("KEBUTUHAN KEGIATAN", width="large", required=True),
                "JUMLAH": st.column_config.NumberColumn("JUMLAH", min_value=0, step=1, width="small"),
                "HARGA SATUAN": st.column_config.NumberColumn("HARGA SATUAN (Rp)", min_value=0, step=1000, format="%d"),
                "HARGA TOTAL": st.column_config.NumberColumn("HARGA TOTAL (Rp)", min_value=0, step=1000, format="%d"),
                "KETERANGAN": st.column_config.TextColumn("KETERANGAN", width="medium"),
            },
            key="editor_table_rab",
        )

        # Reaktif tanpa st.rerun() paksa agar kursor tidak lompat saat pengguna mengetik
        df_rab_live = normalize_items_df(edited_rab_raw, st.session_state.get("ket_tabel_rab", ""))
        total_rab_live = int(df_rab_live["HARGA TOTAL"].sum()) if not df_rab_live.empty else 0

        # Detektor Konsistensi Lokasi Otomatis
        mismatches = detect_location_mismatch(st.session_state.get("lokasi_acara_rab", ""), df_rab_live)
        if mismatches:
            for m in mismatches:
                st.warning(
                    f"⚠️ **Detektor Konsistensi Lokasi**: Judul RAB tertulis **{m['target_city']}**, "
                    f"tetapi Baris #{m['no']} (`{m['item']}`) menyebut **{m['wrong_city']}**."
                )
            if st.button(f"🛠️ Perbaiki Otomatis Nama Kota Menjadi '{mismatches[0]['target_city']}'", use_container_width=False):
                fixed_rows = df_rab_live.to_dict(orient="records")
                for row_f in fixed_rows:
                    for m in mismatches:
                        row_f["KEBUTUHAN KEGIATAN"] = str(row_f["KEBUTUHAN KEGIATAN"]).replace(m["wrong_city"], m["target_city"]).replace(m["wrong_city"].lower(), m["target_city"])
                st.session_state["df_items_rab"] = normalize_items_df(fixed_rows, st.session_state.get("ket_tabel_rab", ""))
                st.rerun()

        st.info(
            f"📌 **Total Item RAB:** `{len(df_rab_live)} Baris`  |  "
            f"💵 **TOTAL ANGGARAN RAB:** `{fmt_rp_cell(total_rab_live, False)}` "
            f"*({terbilang(total_rab_live)})*"
        )
        st.markdown("</div>", unsafe_allow_html=True)

        # CARD 4A & 4B: Catatan Kaki (*Note :) & Kalkulator Sub-Item Note
        cn1, cn2 = st.columns([1.2, 1.1])
        with cn1:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### 📝 CARD 4A: Catatan Kaki RAB (`*Note :` Ber-Bullet)")
            st.text_area(
                "Setiap 1 baris otomatis menjadi 1 poin bullet (•) di bawah tabel RAB:",
                key="catatan_note_rab",
                height=115,
            )
            st.markdown("</div>", unsafe_allow_html=True)

        with cn2:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### 🧮 CARD 4B: Kalkulator Sub-Item `*Note` & Aksi RAB")
            st.caption("Bantu susun rincian gabungan (seperti Iuran/Tol PP) untuk disisipkan ke `*Note :`")
            calc_desc = st.text_input(
                "Ketik Rincian Catatan Baru:",
                placeholder="Contoh: E-Toll (PP) Tol Singosari - Probolinggo Timur Rp 112.000 x 2",
                key="calc_note_input",
            )
            if st.button("➕ Sisipkan Baris ke *Note RAB", use_container_width=True):
                if calc_desc.strip():
                    cur_note = (st.session_state.get("catatan_note_rab") or "").strip()
                    st.session_state["catatan_note_rab"] = (cur_note + NL + calc_desc.strip()).strip()
                    st.rerun()

            ba1, ba2 = st.columns(2)
            with ba1:
                if st.button("💾 Simpan RAB ke Database", use_container_width=True):
                    st.session_state["df_items_rab"] = df_rab_live
                    rab_pay = collect_rab_payload(df_rab_live)
                    ok_s, msg_s = save_rab_to_storage(rab_pay, sheet_id_active, creds_active)
                    if ok_s:
                        st.success(msg_s)
                    else:
                        st.error(msg_s)
            with ba2:
                if st.button("🖨️ Preview & Unduh RAB", use_container_width=True):
                    st.session_state["df_items_rab"] = df_rab_live
                    switch_menu(MENU_PREVIEW, preview_mode="RAB")
            st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================================================
    # VIEW 2: 📝 2. BUAT LPJ (PASCA-KEGIATAN)
    # ==========================================================================
    elif active_menu == MENU_LPJ:
        st.markdown(
            """
            <div class="fily-section-hdr">
                <div>
                    <div style="font-size:17px;font-weight:800;color:#0F172A;">📝 MODUL PEMBUATAN LPJ (PASCA-KEGIATAN)</div>
                    <div style="font-size:12.5px;color:#475569;">Sesuai Pakem Resmi: LPJ_DWP GABUNGAN DI BLITAR_2026-1.docx</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        cl1, cl2 = st.columns([1.2, 1.1])
        with cl1:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### 📌 CARD 1: Identitas Surat & Pelaksanaan Kegiatan LPJ")
            d1, d2 = st.columns([1, 1.6])
            with d1:
                st.date_input("Tanggal Surat LPJ", key="tgl_laporan")
            with d2:
                st.text_input("Nama Kegiatan Utama", key="nama_kegiatan")

            d3, d4 = st.columns(2)
            with d3:
                st.text_input("Tema Kegiatan (Opsional)", key="tema_kegiatan")
            with d4:
                st.text_input("Konteks Tambahan / HUT (Opsional)", key="konteks_hut")

            d5, d6, d7 = st.columns([1, 0.9, 1.2])
            with d5:
                st.date_input("Tanggal Acara", key="tgl_acara")
                st.text_input("Tanggal Manual (Opsional)", key="tgl_acara_manual")
            with d6:
                st.text_input("Waktu Acara", key="waktu_acara")
            with d7:
                st.text_input("Tempat Acara", key="tempat_acara")

            st.text_area("Rangkaian Kegiatan (1 Baris = 1 Nomor Urut Otomatis)", key="rangkaian_kegiatan", height=105)
            st.markdown("</div>", unsafe_allow_html=True)

        with cl2:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### ⚡ CARD 2: Sumber Dana, Tarik dari RAB & Smart Paste WA")
            if st.button("🚀 Tarik Seluruh Data & Tabel dari RAB Aktif Sekarang", use_container_width=True):
                convert_active_rab_to_lpj(df_rab_now)
                st.toast("Data RAB berhasil ditarik ke form LPJ!", icon="🚀")
                st.rerun()

            st.text_input("Sumber Dana (Kalimat Pengantar LPJ)", key="sumber_dana")
            k_col1, k_col2 = st.columns([1.4, 1])
            with k_col1:
                st.text_input("Keterangan Kolom Tabel LPJ (Merged)", key="ket_tabel_dana")
            with k_col2:
                st.write("")
                st.checkbox("Merge Vertikal Keterangan", key="merge_ket_vertikal")

            wa_lpj = st.text_area(
                "Smart Paste Rincian Belanja LPJ dari WhatsApp:",
                height=100,
                placeholder="Sarapan Pagi (Roti) 20 x 13.000" + NL + "Iuran Blitar 10 total 675.000",
                key="wa_paste_lpj",
            )
            bl1, bl2 = st.columns(2)
            with bl1:
                if st.button("🔮 Tambah ke Tabel LPJ", use_container_width=True):
                    new_l = smart_paste_parser(wa_lpj, st.session_state.get("ket_tabel_dana", ""))
                    if new_l:
                        cur = st.session_state["df_items_dana"].to_dict(orient="records")
                        cur.extend(new_l)
                        st.session_state["df_items_dana"] = normalize_items_df(cur, st.session_state.get("ket_tabel_dana", ""))
                        st.rerun()
            with bl2:
                if st.button("🔄 Ganti Seluruh Tabel LPJ", use_container_width=True):
                    new_l = smart_paste_parser(wa_lpj, st.session_state.get("ket_tabel_dana", ""))
                    if new_l:
                        st.session_state["df_items_dana"] = normalize_items_df(new_l, st.session_state.get("ket_tabel_dana", ""))
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        # CARD 3: Tabel Rincian Realisasi Dana LPJ (Anti Cursor-Jump)
        st.markdown('<div class="fily-card">', unsafe_allow_html=True)
        hl1, hl2 = st.columns([2.2, 1])
        with hl1:
            st.markdown("#### 💰 CARD 3: Tabel Realisasi Dana LPJ Interaktif (Logika 3-Mode)")
        with hl2:
            if st.button("🔄 Hitung Ulang & Rapikan Nomor Urut LPJ", use_container_width=True):
                st.session_state["df_items_dana"] = normalize_items_df(
                    st.session_state.get("df_items_dana", pd.DataFrame()),
                    st.session_state.get("ket_tabel_dana", ""),
                )
                st.rerun()

        edited_lpj_raw = st.data_editor(
            st.session_state["df_items_dana"],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "NO": st.column_config.NumberColumn("NO", width="small", disabled=True),
                "KEBUTUHAN KEGIATAN": st.column_config.TextColumn("KEBUTUHAN KEGIATAN", width="large", required=True),
                "JUMLAH": st.column_config.NumberColumn("JUMLAH", min_value=0, step=1, width="small"),
                "HARGA SATUAN": st.column_config.NumberColumn("HARGA SATUAN (Rp)", min_value=0, step=1000, format="%d"),
                "HARGA TOTAL": st.column_config.NumberColumn("HARGA TOTAL (Rp)", min_value=0, step=1000, format="%d"),
                "KETERANGAN": st.column_config.TextColumn("KETERANGAN", width="medium"),
            },
            key="editor_table_lpj",
        )

        df_lpj_live = normalize_items_df(edited_lpj_raw, st.session_state.get("ket_tabel_dana", ""))
        total_lpj_live = int(df_lpj_live["HARGA TOTAL"].sum()) if not df_lpj_live.empty else 0

        st.info(
            f"📌 **Total Item LPJ:** `{len(df_lpj_live)} Baris`  |  "
            f"💵 **TOTAL REALISASI LPJ:** `{fmt_rp_cell(total_lpj_live, False)}` "
            f"*({terbilang(total_lpj_live)})*"
        )
        st.markdown("</div>", unsafe_allow_html=True)

        # CARD 4: Penutup, Tanda Tangan & Simpan LPJ
        cp1, cp2 = st.columns([1.25, 1.05])
        with cp1:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### 📝 CARD 4A: Paragraf Ucapan Donatur & Catatan Instagram")
            st.checkbox("Tampilkan Paragraf Terima Kasih Donatur", key="tampilkan_ucapan_donatur")
            if st.session_state.get("tampilkan_ucapan_donatur", True):
                st.text_area("Redaksi Ucapan Terima Kasih Donatur", key="teks_ucapan_donatur", height=90)
            st.text_input("Catatan Dokumentasi Instagram", key="catatan_ig")
            st.markdown("</div>", unsafe_allow_html=True)

        with cp2:
            st.markdown('<div class="fily-card">', unsafe_allow_html=True)
            st.markdown("#### 🖋️ CARD 4B: Penandatangan & Aksi Simpan LPJ")
            st.text_input("Jabatan Baris 1", key="jabatan_baris_1")
            st.text_input("Jabatan Baris 2", key="jabatan_baris_2")
            st.text_input("Nama Penandatangan", key="nama_penandatangan")
            st.checkbox("Sertakan Gambar Tanda Tangan Asli (Ny. Rini Johan Pandores)", key="tampilkan_ttd")

            al1, al2 = st.columns(2)
            with al1:
                if st.button("💾 Simpan LPJ ke Database", use_container_width=True):
                    st.session_state["df_items_dana"] = df_lpj_live
                    lpj_pay = collect_lpj_payload(df_lpj_live)
                    ok_l, msg_l = save_lpj_to_storage(lpj_pay, sheet_id_active, creds_active)
                    if ok_l:
                        st.success(msg_l)
                    else:
                        st.error(msg_l)
            with al2:
                if st.button("🖨️ Preview & Unduh LPJ", use_container_width=True):
                    st.session_state["df_items_dana"] = df_lpj_live
                    switch_menu(MENU_PREVIEW, preview_mode="LPJ")
            st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================================================
    # VIEW 3: 🖨️ 3. PREVIEW A4 & DOWNLOAD CENTER (SATU-SATUNYA TEMPAT RENDER PDF)
    # ==========================================================================
    elif active_menu == MENU_PREVIEW:
        st.markdown('<div class="fily-card">', unsafe_allow_html=True)
        p_col1, p_col2 = st.columns([1.3, 1.7])
        with p_col1:
            doc_mode_sel = st.radio(
                "Pilih Dokumen yang Ingin Di-Preview & Diunduh:",
                options=["LPJ", "RAB"],
                format_func=lambda x: "📝 Dokumen LPJ Resmi (Pasca-Kegiatan)" if x == "LPJ" else "📋 Dokumen RAB Resmi (Pra-Kegiatan)",
                horizontal=True,
                key="preview_doc_mode",
            )
        with p_col2:
            up_sig_custom = st.file_uploader("Upload Tanda Tangan Baru PNG/JPG (Opsional, khusus LPJ)", type=["png", "jpg", "jpeg"], key="up_sig_preview")
            up_fotos = st.file_uploader("Upload Lampiran Foto/Kwitansi Grid 2x2 (Opsional, khusus LPJ)", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="up_fotos_preview")

        custom_sig_bytes = up_sig_custom.getvalue() if up_sig_custom is not None else None
        attachment_imgs = []
        if up_fotos:
            for idx_f, f_obj in enumerate(up_fotos, start=1):
                attachment_imgs.append({"bytes": f_obj.getvalue(), "caption": f"Dokumentasi {idx_f}: {f_obj.name}"})

        if doc_mode_sel == "RAB":
            active_payload = collect_rab_payload(df_rab_now)
            docx_bytes = generate_rab_docx_bytes(active_payload)
            clean_name = "".join(ch if ch.isalnum() else "_" for ch in active_payload.get("lokasi_acara_rab", "JEMBER")[:25]).strip("_")
            file_base = f"RAB_DWP_{clean_name or 'MALANG'}_{date.today().strftime('%Y%m%d')}"
        else:
            active_payload = collect_lpj_payload(df_lpj_now)
            docx_bytes = generate_lpj_docx_bytes(active_payload, custom_sig_bytes=custom_sig_bytes, attachment_images=attachment_imgs)
            clean_name = "".join(ch if ch.isalnum() else "_" for ch in active_payload.get("tempat_acara", "BLITAR")[:25]).strip("_")
            file_base = f"LPJ_DWP_{clean_name or 'MALANG'}_{date.today().strftime('%Y%m%d')}"

        pdf_bytes, pdf_msg = convert_docx_bytes_to_pdf_bytes(docx_bytes)

        dl1, dl2, dl3 = st.columns([1.1, 1.15, 1.15])
        with dl1:
            if st.button(f"💾 Simpan {doc_mode_sel} Ini ke Google Sheets", use_container_width=True):
                if doc_mode_sel == "RAB":
                    ok_s, msg_s = save_rab_to_storage(active_payload, sheet_id_active, creds_active)
                else:
                    ok_s, msg_s = save_lpj_to_storage(active_payload, sheet_id_active, creds_active)
                if ok_s:
                    st.success(msg_s)
                else:
                    st.error(msg_s)

        with dl2:
            st.markdown(
                make_direct_download_html(
                    file_bytes=docx_bytes,
                    filename=f"{file_base}.docx",
                    mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    label=f"📄 Unduh Word {doc_mode_sel} Resmi (.docx)",
                    bg_gradient="linear-gradient(135deg, #2B98F0 0%, #1D7FE0 100%)",
                ),
                unsafe_allow_html=True,
            )

        with dl3:
            if pdf_bytes:
                st.markdown(
                    make_direct_download_html(
                        file_bytes=pdf_bytes,
                        filename=f"{file_base}.pdf",
                        mime_type="application/pdf",
                        label=f"📕 Unduh PDF {doc_mode_sel} Siap Sebar (.pdf)",
                        bg_gradient="linear-gradient(135deg, #00D09C 0%, #00B386 100%)",
                    ),
                    unsafe_allow_html=True,
                )
            else:
                st.info(pdf_msg)
        st.markdown("</div>", unsafe_allow_html=True)

        a4_html = build_a4_live_preview_html(doc_mode_sel, active_payload, custom_sig_bytes=custom_sig_bytes)
        st.components.v1.html(a4_html, height=1080, scrolling=True)

    # ==========================================================================
    # VIEW 4: 🗄️ 4. DATABASE CLOUD (RIWAYAT RAB, LPJ & REKAP ITEM DANA)
    # ==========================================================================
    elif active_menu == MENU_CLOUD:
        st.markdown('<div class="fily-card">', unsafe_allow_html=True)
        st.markdown("#### 🗄️ Manajemen Arsip Database Google Sheets (`DATABASE_RAB`, `DATABASE_LPJ`, `REKAP_ITEM_DANA`)")
        rab_rows = bundle.get("rab_rows", [])
        lpj_rows = bundle.get("lpj_rows", [])
        rk_rows = bundle.get("rekap_rows", [])

        db_c1, db_c2 = st.columns(2)
        with db_c1:
            st.markdown("##### 📋 Muat / Edit Ulang Arsip RAB")
            if rab_rows:
                rab_map = {f"{r.get('ID_RAB')} | {r.get('LOKASI_ACARA')} ({fmt_rp_cell(r.get('TOTAL_RAB', 0), False)})": r for r in rab_rows}
                sel_rab = st.selectbox("Pilih Arsip RAB:", options=["-- Pilih RAB --"] + list(rab_map.keys()))
                if sel_rab != "-- Pilih RAB --" and st.button("📂 Buka Arsip RAB Terpilih", use_container_width=True):
                    raw_j = rab_map[sel_rab].get("PAYLOAD_JSON", "")
                    if raw_j:
                        load_rab_into_state(json.loads(raw_j))
                        switch_menu(MENU_RAB)
                st.dataframe(pd.DataFrame(rab_rows).drop(columns=["PAYLOAD_JSON"], errors="ignore"), use_container_width=True, hide_index=True)
            else:
                st.info("Belum ada arsip RAB tersimpan.")

        with db_c2:
            st.markdown("##### 📝 Muat / Edit Ulang Arsip LPJ")
            if lpj_rows:
                lpj_map = {f"{r.get('ID_LPJ')} | {r.get('TEMPAT_ACARA')} ({fmt_rp_cell(r.get('TOTAL_DANA', 0), False)})": r for r in lpj_rows}
                sel_lpj = st.selectbox("Pilih Arsip LPJ:", options=["-- Pilih LPJ --"] + list(lpj_map.keys()))
                if sel_lpj != "-- Pilih LPJ --" and st.button("📂 Buka Arsip LPJ Terpilih", use_container_width=True):
                    raw_j = lpj_map[sel_lpj].get("PAYLOAD_JSON", "")
                    if raw_j:
                        load_lpj_into_state(json.loads(raw_j))
                        switch_menu(MENU_LPJ)
                st.dataframe(pd.DataFrame(lpj_rows).drop(columns=["PAYLOAD_JSON"], errors="ignore"), use_container_width=True, hide_index=True)
            else:
                st.info("Belum ada arsip LPJ tersimpan.")

        st.markdown("##### 📊 Tabel Gabungan `REKAP_ITEM_DANA` (Seluruh Baris Item RAB & LPJ)")
        if rk_rows:
            st.dataframe(pd.DataFrame(rk_rows), use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada rincian item tersimpan di REKAP_ITEM_DANA.")
        st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()