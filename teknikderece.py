import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

hisse_listesi = ['A1CAP','ACSEL','ADEL','ADESE','ADGYO','AEFES','AFYON','AGESA','AGHOL','AGROT','AGYO','AHGAZ','AKBNK','AKCNS','AKENR','AKFGY','AKFYE','AKGRT','AKMGY','AKSA','AKSEN','AKSGY','AKSUE','AKYHO','ALARK','ALBRK','ALCAR','ALCTL','ALFAS','ALGYO','ALKA','ALKIM','ALMAD','ALTNY','ALVES','ANELE','ANGEN','ANHYT','ANSGR','ARASE','ARCLK','ARDYZ','ARENA','ARSAN','ARTMS','ARZUM','ASELS','ASGYO','ASTOR','ASUZU','ATAGY','ATAKP','ATATP','ATEKS','ATLAS','ATSYH','AVGYO','AVHOL','AVOD','AVPGY','AVTUR','AYCES','AYDEM','AYEN','AYES','AYGAZ','AZTEK','BAGFS','BAKAB','BALAT','BANVT','BARMA','BASCM','BASGZ','BAYRK','BEGYO','BERA','BEYAZ','BFREN','BIENY','BIGCH','BIMAS','BINHO','BIOEN','BIZIM','BJKAS','BLCYT','BMSCH','BMSTL','BNTAS','BOBET','BORLS','BORSK','BOSSA','BRISA','BRKO','BRKSN','BRKVY','BRLSM','BRMEN','BRSAN','BRYAT','BSOKE','BTCIM','BUCIM','BURCE','BURVA','BVSAN','BYDNR','CANTE','CASA','CATES','CCOLA','CELHA','CEMAS','CEMTS','CEOEM','CIMSA','CLEBI','CMBTN','CMENT','CONSE','COSMO','CRDFA','CRFSA','CUSAN','CVKMD','CWENE','DAGHL','DAGI','DAPGM','DARDL','DENGE','DERHL','DERIM','DESA','DESPC','DEVA','DGATE','DGGYO','DGNMO','DIRIT','DITAS','DMRGD','DMSAS','DNISI','DOAS','DOBUR','DOCO','DOFER','DOGUB','DOHOL','DOKTA','DURDO','DYOBY','DZGYO','EBEBK','ECILC','ECZYT','EDATA','EDIP','EGEEN','EGEPO','EGGUB','EGPRO','EGSER','EKGYO','EKIZ','EKOS','EKSUN','ELITE','EMKEL','EMNIS','ENERY','ENJSA','ENKAI','ENSRI','ENTRA','EPLAS','ERBOS','ERCB','EREGL','ERSU','ESCAR','ESCOM','ESEN','ETILR','ETYAT','EUHOL','EUKYO','EUPWR','EUREN','EUYO','EYGYO','FADE','FENER','FLAP','FMIZP','FONET','FORMT','FORTE','FRIGO','FROTO','FZLGY','GARAN','GARFA','GEDIK','GEDZA','GENIL','GENTS','GEREL','GESAN','GIPTA','GLBMD','GLCVY','GLRYH','GLYHO','GMTAS','GOKNR','GOLTS','GOODY','GOZDE','GRNYO','GRSEL','GRTRK','GSDDE','GSDHO','GSRAY','GUBRF','GWIND','GZNMI','HALKB','HATEK','HATSN','HDFGS','HEDEF','HEKTS','HKTM','HLGYO','HRKET','HTTBT','HUBVC','HUNER','HURGZ','ICBCT','ICUGS','IDGYO','IEYHO','IHAAS','IHEVA','IHGZT','IHLAS','IHLGM','IHYAY','IMASM','INDES','INFO','INGRM','INTEM','INVEO','INVES','IPEKE','ISATR','ISBIR','ISBTR','ISCTR','ISDMR','ISFIN','ISGSY','ISGYO','ISKPL','ISKUR','ISMEN','ISSEN','ISYAT','IZENR','IZFAS','IZINV','IZMDC','JANTS','KAPLM','KAREL','KARSN','KARTN','KARYE','KATMR','KAYSE','KBORU','KCAER','KCHOL','KENT','KERVN','KERVT','KFEIN','KGYO','KIMMR','KLGYO','KLKIM','KLMSN','KLNMA','KLRHO','KLSER','KLSYN','KMPUR','KNFRT','KOCMT','KONKA','KONTR','KONYA','KOPOL','KORDS','KOTON','KOZAA','KOZAL','KRDMA','KRDMB','KRDMD','KRGYO','KRONT','KRPLS','KRSTL','KRTEK','KRVGD','KSTUR','KTLEV','KTSKR','KUTPO','KUVVA','KUYAS','KZBGY','KZGYO','LIDER','LIDFA','LILAK','LINK','LKMNH','LMKDC','LOGO','LRSHO','LUKSK','MAALT','MACKO','MAGEN','MAKIM','MAKTK','MANAS','MARBL','MARKA','MARTI','MAVI','MEDTR','MEGAP','MEGMT','MEKAG','MEPET','MERCN','MERIT','MERKO','METRO','METUR','MGROS','MHRGY','MIATK','MIPAZ','MMCAS','MNDRS','MNDTR','MOBTL','MOGAN','MPARK','MRGYO','MRSHL','MSGYO','MTRKS','MTRYO','MZHLD','NATEN','NETAS','NIBAS','NTGAZ','NTHOL','NUGYO','NUHCM','OBAMS','OBASE','ODAS','ODINE','OFSYM','ONCSM','ORCAY','ORGE','ORMA','OSMEN','OSTIM','OTKAR','OTTO','OYAKC','OYAYO','OYLUM','OYYAT','OZGYO','OZKGY','OZRDN','OZSUB','PAGYO','PAMEL','PAPIL','PARSN','PASEU','PATEK','PCILT','PEGYO','PEKGY','PENGD','PENTA','PETKM','PETUN','PGSUS','PINSU','PKART','PKENT','PLTUR','PNLSN','PNSUT','POLHO','POLTK','PRDGS','PRKAB','PRKME','PRZMA','PSDTC','PSGYO','QNBFB','QNBFL','QUAGR','RALYH','RAYSG','REEDR','RGYAS','RNPOL','RODRG','ROYAL','RTALB','RUBNS','RYGYO','RYSAS','SAFKR','SAHOL','SAMAT','SANEL','SANFM','SANKO','SARKY','SASA','SAYAS','SDTTR','SEGYO','SEKFK','SEKUR','SELEC','SELGD','SELVA','SEYKM','SILVR','SISE','SKBNK','SKTAS','SKYLP','SKYMD','SMART','SMRTG','SNGYO','SNICA','SNKRN','SNPAM','SODSN','SOKE','SOKM','SONME','SRVGY','SUMAS','SUNTK','SURGY','SUWEN','TABGD','TARKM','TATEN','TATGD','TAVHL','TBORG','TCELL','TDGYO','TEKTU','TERA','TETMT','TEZOL','TGSAS','THYAO','TKFEN','TKNSA','TLMAN','TMPOL','TMSN','TNZTP','TOASO','TRCAS','TRGYO','TRILC','TSGYO','TSKB','TSPOR','TTKOM','TTRAK','TUCLK','TUKAS','TUPRS','TUREX','TURGG','TURSG','UFUK','ULAS','ULKER','ULUFA','ULUSE','ULUUN','UMPAS','UNLU','USAK','UZERB','VAKBN','VAKFN','VAKKO','VANGD','VBTYZ','VERTU','VERUS','VESBE','VESTL','VKFYO','VKGYO','VKING','VRGYO','YAPRK','YATAS','YAYLA','YBTAS','YEOTK','YESIL','YGGYO','YGYO','YKBNK','YKSLN','YONGA','YUNSA','YYAPI','YYLGD','ZEDUR','ZOREN','ZRGYO']

durum_sinif_eslestirme = {
    "container-strong-buy": "Güçlü Al",
    "container-buy": "Al",
    "container-neutral": "Nötr",
    "container-sell": "Sat",
    "container-strong-sell": "Güçlü Sat"
}

hisse_durum_listesi = []

for hisse in hisse_listesi:
    url = f"https://tr.tradingview.com/symbols/BIST-{hisse}/technicals/"
    driver.get(url)

    hisse_kodu = hisse

    try:
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.TAG_NAME, 'title')))
    except TimeoutException:
        print(f"Sayfa yüklenirken hata oluştu: {hisse_kodu}")
        hisse_durum_listesi.append((hisse_kodu, "Durum bilgisi bulunamadı"))
        continue

    retries = 3
    for attempt in range(retries):
        try:
            durum_element = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'div.speedometerWrapper-kg4MJrFB.summary-kg4MJrFB div[class*="container-"]'))
            )

            durum_class = durum_element.get_attribute('class')

            durum_text = "Durum bilgisi bulunamadı"
            for class_name, durum in durum_sinif_eslestirme.items():
                if class_name in durum_class:
                    durum_text = durum
                    break


            if durum_text == "Nötr" and attempt < retries - 1:
                print(f"'{hisse_kodu}' için durum 'Nötr' olarak bulundu. Yeniden kontrol ediliyor...")
                time.sleep(2)
                continue

            hisse_durum_listesi.append((hisse_kodu, durum_text))
            break

        except (TimeoutException, NoSuchElementException) as e:
            if attempt < retries - 1:
                print(f"'{hisse_kodu}' için durum bilgisi çekilirken hata oluştu: {e}. Yeniden deneniyor...")
                time.sleep(2)
                continue
            else:
                print(f"'{hisse_kodu}' için durum bilgisi çekilemedi: {e}")
                hisse_durum_listesi.append((hisse_kodu, "Durum bilgisi bulunamadı"))

driver.quit()

df = pd.DataFrame(hisse_durum_listesi, columns=["Hisse Kodu", "Durum"])

excel_dosyasi = "hisse_teknikdurumlari.xlsx"
df.to_excel(excel_dosyasi, index=False)

print(f"Veriler {excel_dosyasi} dosyasına yazdırıldı.")
