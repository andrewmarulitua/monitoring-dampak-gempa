from datetime import datetime
import os
import random
import schedule
import tempfile
import time
from config_se2026 import (
    BASE_PATH,
    LATEST_FILE,  # Anda bisa ganti/sesuaikan dengan path file rekap di config_se2026 atau mendefinisikannya langsung
    NAMA_KABUPATEN,
    archive_filename,
)
import pandas as pd
import requests

# ================= SETTINGS =================
URL_DATA = "https://fasih-sm.bps.go.id/app/api/analytic/api/v1/user-progress/progress-users-by-role"
base_path = BASE_PATH  # FOLDER UNTUK MENYIMPAN DATA HASIL SCRAPPING

# Tentukan nama file Excel utama gabungan
MAIN_EXCEL_OUTPUT = os.path.join(base_path, "SCRAPING_REKAP_DAMPAK_BENCANA_ENDE_LATEST.xlsx")
# ==========================================================


# ===================== CONFIGURATIONS (2 TARGET SCRAPING) =====================
SCRAPE_CONFIGS = [
    {
        "name": "Keluarga",
        "sheet_name": "Keluarga",
        "cookies": {
            'f5avraaaaaaaaaaaaaaaa_session_': 'BNLNBJGIPGPMCPHHKIEMHOAPFJFGJHOJFNFIFOCMBJNFBFGKMBBBKPDMAFJAGOIOLHBDJOAJPIJPKHPGAHNACFGDOCJLPDPNMDOCCJPJIGDKIIGOGHEMFBANDJJNNAOJ',
            'f5avraaaaaaaaaaaaaaaa_session_': 'FKLCMDHNELFIJLBGICOKKLAGINBIOICALGPFCBBKPDPMLLBPEGKKAIBANGLPAJLGHDHDAJOGBOGADLEMDFOAIMEKBAGGNAKGODEDGEDLKNNGKIHKBFEPFLIANONECGDL',
            'f5_cspm': '1234',
            '_ga_ER466FVWPF': 'GS2.1.s1791523792$o1$g0$t1791523807$j45$l0$h0',
            '_ga': 'GA1.3.121889234.1791523793',
            '_ga_XXTTVXWHDB': 'GS2.3.s1791526012$o2$g0$t1791526012$j60$l0$h0',
            'cf_clearance': 'p9_LSDdK8unlQElPJ8kqX.78NhgqC19FEcokxooUWHY-1792806076-1.2.1.1-zDSaIM6ge7QnlQ24omEK4b6VZJgxRsOZui4SSRFJitsb8S8is4naaN13HEpiEyfzZSj32pDOW2Oiq6NBMHlIRM0qGid0Dq0L7JyW9pBBdC1Btsc9NY5a0PGDpweHykIPacC3WRTzx6qQCmw_vItlbDrFinbvp5NwV8f2mZYoW0J.OxDhDHHcbSenTNpk2hJOe6Eu6wQefJWJbldG8DCA54Jycf6SXzfY7wRvJSMAjx5QI7dahSdhaa0tCGlgE68cWUOf3P5sN.sjvNmbIl7otncQkE5d4FfPY4yFAPRnA6_Y3IL3SWxVnFH7QC01mXz8jy3qgl68VBrHe3Zjg3JLCTWPW9yVkzBQ7sC1AqcomroKrOguC.j2nVavXgdHSP.Qc5cankgxtxAh57cifpwiSjCKXqje0FhTt4NLLaUPvvC12EyI5ebisjnvmA_gwK0kzUs5lSRH5YL85l2ZHVIORNHAzh905Nh5qIrh13HLMTk',
            'XSRF-TOKEN': 'c9c60815-733e-40ff-82de-eb39b39184f8',
            'JSESSIONID': '9BECACF6AD8CFCAF418C098313C4539D',
            'SESSION': '297e4445-bc21-446b-8a0d-8f2af5a84518',
            'f5avraaaaaaaaaaaaaaaa_session_': 'DKPKNPIMEDCFBPPAODNFMKFNJKENHEMNHOHIHPEAMNPBJFBFIMBODFMJBMBIPHFMMIFDJHLMNGCAADPGIEDALLJLNBELAHKHHJECDPICJNIEFILMBGNOMBLJCHIJFHCD',
            'TS01433fd3': '01266d26d0ff3aa048f9de3ea50d6a6f38d12898e15625c2d1e9f1548c3f616bada5d365f846e5537a466be6559c1304049a14d565',
            'TS00000000076': '0868f8be6fab28007143b64786293536475820d8a4490c0430a9cc07c585a5ffa14b7019a8d85e316849b726bb8df61608b4c8f5a409d0006417b914c1cf75e75420db6a10ac5e1ee9ce02cd708ef740279a11b5f694b4f4ac653f185c2b883db2a2df077d3d26fbd28dfa2371f84c26b09ce47886e64038a8c07ba8b3982f017d1397e0e9cef6364172552916c41a2caf682d36e12730b10e03bf30d111eda1670324bbc605b1dea5f90a3385d69e73a0461c6290999984a437e8ff2a17100439b0df8e4a0cb3ee2b1eb1abd0ecf70cdc4e558555a1e2c53231c8c727ac44f354a6972669f86e53b84abd1ac3a2e8119abd7b6663baa0a333f8f1e6dc586bd985606fdd08b9e9fb',
            'TSPD_101_DID': '0868f8be6fab28007143b64786293536475820d8a4490c0430a9cc07c585a5ffa14b7019a8d85e316849b726bb8df61608b4c8f5a4063800a70d093e910adbbbb6e48400996af5151bf3caf803a6273f8a37471c111608e730327bf5b3a1133e49715392f6c7b2cfba4d37cceef0d847',
            'TSPD_101': '0868f8be6fab2800e8aca4b83163d916d19840350351d40c7b9bee20d082910cda9a11a42150f3709ffe5cc66e3b0f360841ffe7cc0518009e8d0d09aff128dd5ca1732140a3428bba23ce13beb1c95e',
            'TS011f2d1a': '01266d26d0057b949f5ef51e5f8ea11e8a8449d353c5881cf45b9e88d6962a928d17f1fac4de47c3d9c6250c8741c1d6651b25ff01',
            'TS5220f739077': '0868f8be6fab280068adcd1fc8f7fbf2adc1b07c29b019083eae45adce4964e671eb938acf136b061b554b59b187b9b5083ed0848417200013880b71b65867d0cd97bf84f94ad31bc71f570a894e6f454963045c5daaa74c',
            'TS5220f739029': '0868f8be6fab2800e1cf1edd3358674f53865576227afd24c6605025455c27007bc20cf790dd81188b11e11cfe682ab2',
            'TSf1edb2d2027': '0868f8be6fab2000d6dc4be35edafc558142eabaf9c32045cdd8b17096b9a97253daf646001c3c7008c84ef16b1130005ae074065c55d1f0ba757432ddb5f9999bf3c78330f945ab1889818ad6e0af894e47b7c3f3e865e4049ad87cf9cb1c93',
        },
        "headers": {
            'Accept': '*/*',
            'Accept-Language': 'en-US,en;q=0.9,id;q=0.8',
            'Connection': 'keep-alive',
            'Content-Type': 'application/json',
            'Origin': 'https://fasih-sm.bps.go.id',
            'Referer': 'https://fasih-sm.bps.go.id/app/surveys/3aaeb894-0fdd-429b-a083-53ddd979f413/e511c9a7-275d-4ce7-9776-8e65f6b4aedd',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Linux; Android 16; Pixel 10) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/156.0.0.0 Mobile Safari/537.36',
            'X-XSRF-TOKEN': 'c9c60815-733e-40ff-82de-eb39b39184f8',
            'sec-ch-ua': '"Not:A-Brand";v="8", "Chromium";v="156", "Google Chrome";v="156"',
            'sec-ch-ua-mobile': '?1',
            'sec-ch-ua-platform': '"Android"',
            'Cookie': 'f5avraaaaaaaaaaaaaaaa_session_=BNLNBJGIPGPMCPHHKIEMHOAPFJFGJHOJFNFIFOCMBJNFBFGKMBBBKPDMAFJAGOIOLHBDJOAJPIJPKHPGAHNACFGDOCJLPDPNMDOCCJPJIGDKIIGOGHEMFBANDJJNNAOJ; f5avraaaaaaaaaaaaaaaa_session_=FKLCMDHNELFIJLBGICOKKLAGINBIOICALGPFCBBKPDPMLLBPEGKKAIBANGLPAJLGHDHDAJOGBOGADLEMDFOAIMEKBAGGNAKGODEDGEDLKNNGKIHKBFEPFLIANONECGDL; f5_cspm=1234; _ga_ER466FVWPF=GS2.1.s1791523792$o1$g0$t1791523807$j45$l0$h0; _ga=GA1.3.121889234.1791523793; _ga_XXTTVXWHDB=GS2.3.s1791526012$o2$g0$t1791526012$j60$l0$h0; cf_clearance=p9_LSDdK8unlQElPJ8kqX.78NhgqC19FEcokxooUWHY-1792806076-1.2.1.1-zDSaIM6ge7QnlQ24omEK4b6VZJgxRsOZui4SSRFJitsb8S8is4naaN13HEpiEyfzZSj32pDOW2Oiq6NBMHlIRM0qGid0Dq0L7JyW9pBBdC1Btsc9NY5a0PGDpweHykIPacC3WRTzx6qQCmw_vItlbDrFinbvp5NwV8f2mZYoW0J.OxDhDHHcbSenTNpk2hJOe6Eu6wQefJWJbldG8DCA54Jycf6SXzfY7wRvJSMAjx5QI7dahSdhaa0tCGlgE68cWUOf3P5sN.sjvNmbIl7otncQkE5d4FfPY4yFAPRnA6_Y3IL3SWxVnFH7QC01mXz8jy3qgl68VBrHe3Zjg3JLCTWPW9yVkzBQ7sC1AqcomroKrOguC.j2nVavXgdHSP.Qc5cankgxtxAh57cifpwiSjCKXqje0FhTt4NLLaUPvvC12EyI5ebisjnvmA_gwK0kzUs5lSRH5YL85l2ZHVIORNHAzh905Nh5qIrh13HLMTk; XSRF-TOKEN=c9c60815-733e-40ff-82de-eb39b39184f8; JSESSIONID=9BECACF6AD8CFCAF418C098313C4539D; SESSION=297e4445-bc21-446b-8a0d-8f2af5a84518; f5avraaaaaaaaaaaaaaaa_session_=DKPKNPIMEDCFBPPAODNFMKFNJKENHEMNHOHIHPEAMNPBJFBFIMBODFMJBMBIPHFMMIFDJHLMNGCAADPGIEDALLJLNBELAHKHHJECDPICJNIEFILMBGNOMBLJCHIJFHCD; TS01433fd3=01266d26d0ff3aa048f9de3ea50d6a6f38d12898e15625c2d1e9f1548c3f616bada5d365f846e5537a466be6559c1304049a14d565; TS00000000076=0868f8be6fab28007143b64786293536475820d8a4490c0430a9cc07c585a5ffa14b7019a8d85e316849b726bb8df61608b4c8f5a409d0006417b914c1cf75e75420db6a10ac5e1ee9ce02cd708ef740279a11b5f694b4f4ac653f185c2b883db2a2df077d3d26fbd28dfa2371f84c26b09ce47886e64038a8c07ba8b3982f017d1397e0e9cef6364172552916c41a2caf682d36e12730b10e03bf30d111eda1670324bbc605b1dea5f90a3385d69e73a0461c6290999984a437e8ff2a17100439b0df8e4a0cb3ee2b1eb1abd0ecf70cdc4e558555a1e2c53231c8c727ac44f354a6972669f86e53b84abd1ac3a2e8119abd7b6663baa0a333f8f1e6dc586bd985606fdd08b9e9fb; TSPD_101_DID=0868f8be6fab28007143b64786293536475820d8a4490c0430a9cc07c585a5ffa14b7019a8d85e316849b726bb8df61608b4c8f5a4063800a70d093e910adbbbb6e48400996af5151bf3caf803a6273f8a37471c111608e730327bf5b3a1133e49715392f6c7b2cfba4d37cceef0d847; TSPD_101=0868f8be6fab2800e8aca4b83163d916d19840350351d40c7b9bee20d082910cda9a11a42150f3709ffe5cc66e3b0f360841ffe7cc0518009e8d0d09aff128dd5ca1732140a3428bba23ce13beb1c95e; TS011f2d1a=01266d26d0057b949f5ef51e5f8ea11e8a8449d353c5881cf45b9e88d6962a928d17f1fac4de47c3d9c6250c8741c1d6651b25ff01; TS5220f739077=0868f8be6fab280068adcd1fc8f7fbf2adc1b07c29b019083eae45adce4964e671eb938acf136b061b554b59b187b9b5083ed0848417200013880b71b65867d0cd97bf84f94ad31bc71f570a894e6f454963045c5daaa74c; TS5220f739029=0868f8be6fab2800e1cf1edd3358674f53865576227afd24c6605025455c27007bc20cf790dd81188b11e11cfe682ab2; TSf1edb2d2027=0868f8be6fab2000d6dc4be35edafc558142eabaf9c32045cdd8b17096b9a97253daf646001c3c7008c84ef16b1130005ae074065c55d1f0ba757432ddb5f9999bf3c78330f945ab1889818ad6e0af894e47b7c3f3e865e4049ad87cf9cb1c93',
        },
        "json_data": {
            'surveyPeriodId': 'e511c9a7-275d-4ce7-9776-8e65f6b4aedd',
            'surveyRoleId': '95964fb3-f998-4190-9cf4-6c2a0e77a44c',
            'size': 10,
            'page': 0,
            'search': '',
            'target': 'TARGET_ONLY',
            'region': {
                'region1Id': None,
                'region2Id': None,
                'region3Id': None,
                'region4Id': None,
                'region5Id': None,
                'region6Id': None,
                'region7Id': None,
                'region8Id': None,
                'region9Id': None,
                'region10Id': None,
            },
            'regionSummaryLevel': 6,
        }
    },
    {
        "name": "Infrastruktur",
        "sheet_name": "Infrastruktur",
        "cookies": {
            'f5avraaaaaaaaaaaaaaaa_session_': 'FMOAJOPFOGJMIHBMMHFNJDAHMOLHKAFJAGHLHMKODCHALNCLIBOHACIKMEGHFJLGJLFDEGJOLHKBCCBIFNPABMIANCEPBCCJFHOHNEAHIBCJMFDCJKCBIGMHFCNKPDEA',
            'f5avraaaaaaaaaaaaaaaa_session_': 'FKLCMDHNELFIJLBGICOKKLAGINBIOICALGPFCBBKPDPMLLBPEGKKAIBANGLPAJLGHDHDAJOGBOGADLEMDFOAIMEKBAGGNAKGODEDGEDLKNNGKIHKBFEPFLIANONECGDL',
            'f5_cspm': '1234',
            '_ga_ER466FVWPF': 'GS2.1.s1791523792$o1$g0$t1791523807$j45$l0$h0',
            '_ga': 'GA1.3.121889234.1791523793',
            '_ga_XXTTVXWHDB': 'GS2.3.s1791526012$o2$g0$t1791526012$j60$l0$h0',
            'cf_clearance': 'p9_LSDdK8unlQElPJ8kqX.78NhgqC19FEcokxooUWHY-1792806076-1.2.1.1-zDSaIM6ge7QnlQ24omEK4b6VZJgxRsOZui4SSRFJitsb8S8is4naaN13HEpiEyfzZSj32pDOW2Oiq6NBMHlIRM0qGid0Dq0L7JyW9pBBdC1Btsc9NY5a0PGDpweHykIPacC3WRTzx6qQCmw_vItlbDrFinbvp5NwV8f2mZYoW0J.OxDhDHHcbSenTNpk2hJOe6Eu6wQefJWJbldG8DCA54Jycf6SXzfY7wRvJSMAjx5QI7dahSdhaa0tCGlgE68cWUOf3P5sN.sjvNmbIl7otncQkE5d4FfPY4yFAPRnA6_Y3IL3SWxVnFH7QC01mXz8jy3qgl68VBrHe3Zjg3JLCTWPW9yVkzBQ7sC1AqcomroKrOguC.j2nVavXgdHSP.Qc5cankgxtxAh57cifpwiSjCKXqje0FhTt4NLLaUPvvC12EyI5ebisjnvmA_gwK0kzUs5lSRH5YL85l2ZHVIORNHAzh905Nh5qIrh13HLMTk',
            'XSRF-TOKEN': 'c9c60815-733e-40ff-82de-eb39b39184f8',
            'JSESSIONID': '9BECACF6AD8CFCAF418C098313C4539D',
            'SESSION': '297e4445-bc21-446b-8a0d-8f2af5a84518',
            'f5avraaaaaaaaaaaaaaaa_session_': 'DKPKNPIMEDCFBPPAODNFMKFNJKENHEMNHOHIHPEAMNPBJFBFIMBODFMJBMBIPHFMMIFDJHLMNGCAADPGIEDALLJLNBELAHKHHJECDPICJNIEFILMBGNOMBLJCHIJFHCD',
            'TS01433fd3': '01266d26d0ff3aa048f9de3ea50d6a6f38d12898e15625c2d1e9f1548c3f616bada5d365f846e5537a466be6559c1304049a14d565',
            'TS011f2d1a': '01266d26d093b50d1d72d5841bc1fd182a1601d10bf0d1c6957e82991a4ca76a676846fd132a6620647777ed5a504e74cc608bb0bb',
            'TS00000000076': '0868f8be6fab28007143b64786293536475820d8a4490c0430a9cc07c585a5ffa14b7019a8d85e316849b726bb8df61608b4c8f5a409d0006417b914c1cf75e75420db6a10ac5e1ee9ce02cd708ef740279a11b5f694b4f4ac653f185c2b883db2a2df077d3d26fbd28dfa2371f84c26b09ce47886e64038a8c07ba8b3982f017d1397e0e9cef6364172552916c41a2caf682d36e12730b10e03bf30d111eda1670324bbc605b1dea5f90a3385d69e73a0461c6290999984a437e8ff2a17100439b0df8e4a0cb3ee2b1eb1abd0ecf70cdc4e558555a1e2c53231c8c727ac44f354a6972669f86e53b84abd1ac3a2e8119abd7b6663baa0a333f8f1e6dc586bd985606fdd08b9e9fb',
            'TSPD_101_DID': '0868f8be6fab28007143b64786293536475820d8a4490c0430a9cc07c585a5ffa14b7019a8d85e316849b726bb8df61608b4c8f5a4063800a70d093e910adbbbb6e48400996af5151bf3caf803a6273f8a37471c111608e730327bf5b3a1133e49715392f6c7b2cfba4d37cceef0d847',
            'TSPD_101': '0868f8be6fab2800e8aca4b83163d916d19840350351d40c7b9bee20d082910cda9a11a42150f3709ffe5cc66e3b0f360841ffe7cc0518009e8d0d09aff128dd5ca1732140a3428bba23ce13beb1c95e',
            'TS5220f739077': '0868f8be6fab2800f9bd322a0fd6b76a9a9b9341e8a6d52d074c09650b6714674a6bcc043feeb73ea580d79c093ceea008a01287af172000edec879927271ebdcd97bf84f94ad31b698eb5df35ab583a83def985eac66fcc',
            'TS5220f739029': '0868f8be6fab280058cf04490f0d6ab4e0230879f15111880f6827ade3889d1db805398af1e3e16eb9bbaf67d2cb45f0',
            'TSf1edb2d2027': '0868f8be6fab2000a71fba1e1cd5433c4a3161cf94db781a8ff3f089c451b3b2b7e01393d9c3100e08c4e13aa6113000749c9f69749844f061e706fc5544ce9525dbd83772ecec08dd1a29d06fba049cd111f3a2bbcfb0a0b73a1dc862168b4c',
        },
        "headers": {
            'Accept': '*/*',
            'Accept-Language': 'en-US,en;q=0.9,id;q=0.8',
            'Connection': 'keep-alive',
            'Content-Type': 'application/json',
            'Origin': 'https://fasih-sm.bps.go.id',
            'Referer': 'https://fasih-sm.bps.go.id/app/surveys/edb5c502-7062-422e-bcb2-493418b6bee0/860befa8-a519-4571-b79d-4fd745d9f0a8',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Linux; Android 16; Pixel 10) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/156.0.0.0 Mobile Safari/537.36',
            'X-XSRF-TOKEN': 'c9c60815-733e-40ff-82de-eb39b39184f8',
            'sec-ch-ua': '"Not:A-Brand";v="8", "Chromium";v="156", "Google Chrome";v="156"',
            'sec-ch-ua-mobile': '?1',
            'sec-ch-ua-platform': '"Android"',
            'Cookie': 'f5avraaaaaaaaaaaaaaaa_session_=FMOAJOPFOGJMIHBMMHFNJDAHMOLHKAFJAGHLHMKODCHALNCLIBOHACIKMEGHFJLGJLFDEGJOLHKBCCBIFNPABMIANCEPBCCJFHOHNEAHIBCJMFDCJKCBIGMHFCNKPDEA; f5avraaaaaaaaaaaaaaaa_session_=FKLCMDHNELFIJLBGICOKKLAGINBIOICALGPFCBBKPDPMLLBPEGKKAIBANGLPAJLGHDHDAJOGBOGADLEMDFOAIMEKBAGGNAKGODEDGEDLKNNGKIHKBFEPFLIANONECGDL; f5_cspm=1234; _ga_ER466FVWPF=GS2.1.s1791523792$o1$g0$t1791523807$j45$l0$h0; _ga=GA1.3.121889234.1791523793; _ga_XXTTVXWHDB=GS2.3.s1791526012$o2$g0$t1791526012$j60$l0$h0; cf_clearance=p9_LSDdK8unlQElPJ8kqX.78NhgqC19FEcokxooUWHY-1792806076-1.2.1.1-zDSaIM6ge7QnlQ24omEK4b6VZJgxRsOZui4SSRFJitsb8S8is4naaN13HEpiEyfzZSj32pDOW2Oiq6NBMHlIRM0qGid0Dq0L7JyW9pBBdC1Btsc9NY5a0PGDpweHykIPacC3WRTzx6qQCmw_vItlbDrFinbvp5NwV8f2mZYoW0J.OxDhDHHcbSenTNpk2hJOe6Eu6wQefJWJbldG8DCA54Jycf6SXzfY7wRvJSMAjx5QI7dahSdhaa0tCGlgE68cWUOf3P5sN.sjvNmbIl7otncQkE5d4FfPY4yFAPRnA6_Y3IL3SWxVnFH7QC01mXz8jy3qgl68VBrHe3Zjg3JLCTWPW9yVkzBQ7sC1AqcomroKrOguC.j2nVavXgdHSP.Qc5cankgxtxAh57cifpwiSjCKXqje0FhTt4NLLaUPvvC12EyI5ebisjnvmA_gwK0kzUs5lSRH5YL85l2ZHVIORNHAzh905Nh5qIrh13HLMTk; XSRF-TOKEN=c9c60815-733e-40ff-82de-eb39b39184f8; JSESSIONID=9BECACF6AD8CFCAF418C098313C4539D; SESSION=297e4445-bc21-446b-8a0d-8f2af5a84518; f5avraaaaaaaaaaaaaaaa_session_=DKPKNPIMEDCFBPPAODNFMKFNJKENHEMNHOHIHPEAMNPBJFBFIMBODFMJBMBIPHFMMIFDJHLMNGCAADPGIEDALLJLNBELAHKHHJECDPICJNIEFILMBGNOMBLJCHIJFHCD; TS01433fd3=01266d26d0ff3aa048f9de3ea50d6a6f38d12898e15625c2d1e9f1548c3f616bada5d365f846e5537a466be6559c1304049a14d565; TS011f2d1a=01266d26d093b50d1d72d5841bc1fd182a1601d10bf0d1c6957e82991a4ca76a676846fd132a6620647777ed5a504e74cc608bb0bb; TS00000000076=0868f8be6fab28007143b64786293536475820d8a4490c0430a9cc07c585a5ffa14b7019a8d85e316849b726bb8df61608b4c8f5a409d0006417b914c1cf75e75420db6a10ac5e1ee9ce02cd708ef740279a11b5f694b4f4ac653f185c2b883db2a2df077d3d26fbd28dfa2371f84c26b09ce47886e64038a8c07ba8b3982f017d1397e0e9cef6364172552916c41a2caf682d36e12730b10e03bf30d111eda1670324bbc605b1dea5f90a3385d69e73a0461c6290999984a437e8ff2a17100439b0df8e4a0cb3ee2b1eb1abd0ecf70cdc4e558555a1e2c53231c8c727ac44f354a6972669f86e53b84abd1ac3a2e8119abd7b6663baa0a333f8f1e6dc586bd985606fdd08b9e9fb; TSPD_101_DID=0868f8be6fab28007143b64786293536475820d8a4490c0430a9cc07c585a5ffa14b7019a8d85e316849b726bb8df61608b4c8f5a4063800a70d093e910adbbbb6e48400996af5151bf3caf803a6273f8a37471c111608e730327bf5b3a1133e49715392f6c7b2cfba4d37cceef0d847; TSPD_101=0868f8be6fab2800e8aca4b83163d916d19840350351d40c7b9bee20d082910cda9a11a42150f3709ffe5cc66e3b0f360841ffe7cc0518009e8d0d09aff128dd5ca1732140a3428bba23ce13beb1c95e; TS5220f739077=0868f8be6fab2800f9bd322a0fd6b76a9a9b9341e8a6d52d074c09650b6714674a6bcc043feeb73ea580d79c093ceea008a01287af172000edec879927271ebdcd97bf84f94ad31b698eb5df35ab583a83def985eac66fcc; TS5220f739029=0868f8be6fab280058cf04490f0d6ab4e0230879f15111880f6827ade3889d1db805398af1e3e16eb9bbaf67d2cb45f0; TSf1edb2d2027=0868f8be6fab2000a71fba1e1cd5433c4a3161cf94db781a8ff3f089c451b3b2b7e01393d9c3100e08c4e13aa6113000749c9f69749844f061e706fc5544ce9525dbd83772ecec08dd1a29d06fba049cd111f3a2bbcfb0a0b73a1dc862168b4c',
        },
        "json_data": {
            'surveyPeriodId': '860befa8-a519-4571-b79d-4fd745d9f0a8',
            'surveyRoleId': '867fb5b1-73bf-461b-a936-671acabe807e',
            'size': 10,
            'page': 0,
            'search': '',
            'target': 'TARGET_ONLY',
            'region': {
                'region1Id': None,
                'region2Id': None,
                'region3Id': None,
                'region4Id': None,
                'region5Id': None,
                'region6Id': None,
                'region7Id': None,
                'region8Id': None,
                'region9Id': None,
                'region10Id': None,
            },
            'regionSummaryLevel': 4,
        }
    }
]

# ================================================================

if not os.path.exists(base_path):
    os.makedirs(base_path)


def save_and_merge(new_data, sheet_name="Keluarga"):
    """Menerima data hasil scraping, memperkaya dengan master_data (sesuai sheet),

    lalu memperbarui sheet spesifik di dalam file MAIN_EXCEL_OUTPUT secara aman.
    """
    if not new_data:
        return

    print(f"\n-------------------------------------------------------------")
    print(f"🔄 Memproses dan menyelaraskan data scraping untuk sheet: {sheet_name}...")

    df_target = pd.DataFrame(new_data)
    df_target["scraped_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Baca file master jika tersedia berdasarkan sheet_name
    master_file_path = os.path.join(base_path, "master_data.xlsx")
    if os.path.exists(master_file_path):
        try:
            df_master = pd.read_excel(master_file_path, sheet_name=sheet_name)
        except Exception as e:
            print(f"⚠️ Warning: Gagal membaca sheet '{sheet_name}' dari {master_file_path}: {e}")
            df_master = None

        if df_master is not None:
            columns_to_update = [
                "nmkab", "nmkec", "nmdesa", "nmsls", "nmsubsls",
                "pengawas", "pencacah", "nama_pcl", "nama_pml",
            ]

            df_target["regionCode"] = df_target["regionCode"].astype(str).str.strip()
            df_master["regionCode"] = df_master["regionCode"].astype(str).str.strip()

            available_master_cols = [col for col in columns_to_update if col in df_master.columns]
            df_master_subset = df_master[["regionCode"] + available_master_cols].drop_duplicates(subset=["regionCode"])

            df_merged = pd.merge(
                df_target,
                df_master_subset,
                on="regionCode",
                how="left",
                suffixes=("", "_master"),
            )

            strict_overwrite_cols = ["pengawas", "nama_pml"]

            for col in available_master_cols:
                master_col_name = f"{col}_master"

                if master_col_name in df_merged.columns:
                    if col not in df_merged.columns:
                        df_merged[col] = df_merged[master_col_name]
                    else:
                        if col in strict_overwrite_cols:
                            df_merged[col] = df_merged[master_col_name].fillna(df_merged[col])
                        else:
                            df_merged[col] = df_merged[col].fillna(df_merged[master_col_name])

                    df_merged.drop(columns=[master_col_name], inplace=True)

            if "username" in df_merged.columns and "pencacah" in df_merged.columns:
                mask_mismatch = (
                    df_merged["username"].astype(str).str.strip()
                    != df_merged["pencacah"].astype(str).str.strip()
                )
                df_merged.loc[mask_mismatch, "pencacah"] = df_merged.loc[mask_mismatch, "username"]

            if "pencacah" in df_master.columns and "nama_pcl" in df_master.columns:
                df_master_clean = df_master.dropna(subset=["pencacah", "nama_pcl"]).drop_duplicates(subset=["pencacah"])
                pcl_map = dict(
                    zip(
                        df_master_clean["pencacah"].astype(str).str.strip(),
                        df_master_clean["nama_pcl"],
                    )
                )

                if "pencacah" in df_merged.columns:
                    existing_nama_pcl = df_merged.get("nama_pcl", pd.Series([None] * len(df_merged)))
                    df_merged["nama_pcl"] = (
                        df_merged["pencacah"]
                        .astype(str)
                        .str.strip()
                        .map(pcl_map)
                        .fillna(existing_nama_pcl)
                    )

            df_final = df_merged
        else:
            df_final = df_target
    else:
        print(f"⚠️ Warning: File '{master_file_path}' tidak ditemukan. Menggunakan data tanpa enrichment master.")
        df_final = df_target

    # Konversi Kolom Numerik
    numeric_cols = [
        "total_data", "APPROVED BY Pengawas", "SUBMITTED BY Pencacah", "OPEN",
        "REJECTED BY Pengawas", "DRAFT", "EDITED BY Pengawas", "REVOKED BY Pengawas",
        "SUBMITTED RESPONDENT",
    ]

    for col in numeric_cols:
        if col in df_final.columns:
            df_final[col] = pd.to_numeric(df_final[col], errors="coerce")

    # Simpan Arsip histori terpisah (per scraping)
    safe_sheet_name = sheet_name.lower().replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = archive_filename(f"{safe_sheet_name}_{timestamp}")
    df_final.to_excel(backup_file, index=False)

    # **Simpan & Gabungkan ke File Utama Multi-Sheet secara Atomic**
    _atomic_write_multi_sheet(df_final, sheet_name, MAIN_EXCEL_OUTPUT)
    print(f"💾 Data berhasil diperbarui ke sheet '{sheet_name}' di dalam file: {MAIN_EXCEL_OUTPUT}")
    print("-------------------------------------------------------------\n")


def _atomic_write_multi_sheet(df_new, sheet_name, path):
    """Menulis atau memperbarui sheet tertentu di dalam file Excel yang sama

    tanpa merusak/menghapus sheet lainnya (misal: sheet Keluarga tetap aman saat sheet Infrastruktur ditulis).
    """
    folder = os.path.dirname(path) or "."
    
    # Baca seluruh sheet yang sudah ada jika file utama sudah ada
    existing_sheets = {}
    if os.path.exists(path):
        try:
            existing_sheets = pd.read_excel(path, sheet_name=None)
        except Exception:
            pass

    # Perbarui atau tambahkan sheet yang sedang diproses
    existing_sheets[sheet_name] = df_new

    # Tulis ke file sementara terlebih dahulu (Atomic Write)
    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir=folder)
    os.close(fd)
    
    try:
        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            for s_name, s_df in existing_sheets.items():
                s_df.to_excel(writer, sheet_name=s_name, index=False)
        os.replace(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def auto_push_github():
    import subprocess

    try:
        subprocess.run(["git", "add", "data/"], check=True)
        status = subprocess.run(
            ["git", "status", "--porcelain"], capture_output=True, text=True
        )

        if not status.stdout.strip():
            print("📌 Tidak ada perubahan untuk dipush ke GitHub.")
            return

        subprocess.run(["git", "commit", "-m", "update"], check=True)
        subprocess.run(["git", "pull", "--rebase", "origin", "main"], check=True)
        subprocess.run(["git", "push", "origin", "main"], check=True)

        print("✅ Data berhasil dipush ke GitHub dengan pesan commit 'update'")

    except Exception as e:
        print(f"❌ Error push GitHub: {e}")


def fetch_data(config):
    target_name = config.get("name")
    sheet_name = config.get("sheet_name")
    cookies = config.get("cookies")
    headers = config.get("headers")
    json_payload = config.get("json_data").copy()

    print(f"\n🚀 Memulai scraping untuk Target: [{target_name}]")

    all_rows = []
    page = 0
    size = 10

    while True:
        json_payload["page"] = page
        json_payload["size"] = size

        response = requests.post(
            URL_DATA,
            cookies=cookies,
            headers=headers,
            json=json_payload,
        )

        if response.status_code != 200:
            print(f"❌ Error [{target_name}] di page {page}")
            print(f"Status Code: {response.status_code}")
            print(response.text[:1000])
            break

        json_res = response.json()
        data_block = json_res.get("data", {})
        data = data_block.get("content", [])
        is_last = data_block.get("last", True)

        print(f"📄 [{target_name}] Page {page} | jumlah data: {len(data)} | last: {is_last}")

        for user in data:
            for region in user.get("regionSummary", []):
                row = {
                    "userId": user.get("userId"),
                    "username": user.get("username"),
                    "email": user.get("email"),
                    "role": user.get("roleName"),
                    "regionCode": region.get("regionCode"),
                    "total_data": region.get("total"),
                }

                for status in region.get("statusBreakdown", []):
                    row[status.get("status")] = status.get("count")

                all_rows.append(row)

        if is_last:
            print(f"✅ [{target_name}] Sudah sampai halaman terakhir")
            break

        page += 1
        time.sleep(random.uniform(1, 3))

    if all_rows:
        save_and_merge(all_rows, sheet_name=sheet_name)

    print(f"🎉 Scraping [{target_name}] selesai!")


def job():
    print(
        f"\n[+] Memulai proses scraping berkala pada"
        f" {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    
    for config in SCRAPE_CONFIGS:
        fetch_data(config)
        time.sleep(2)
        
    auto_push_github()


if __name__ == "__main__":
    schedule.every(4).hours.do(job)

    print(
        "⏱️  Script berjalan otomatis setiap 4 jam. Tekan Ctrl+C untuk menghentikan."
    )

    job()

    while True:
        schedule.run_pending()
        time.sleep(1)