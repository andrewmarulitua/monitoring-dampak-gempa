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
            'f5avraaaaaaaaaaaaaaaa_session_': 'BHKPGAHHMKMBOINLLLPMLGHMDNFLPEOPMEKMBKLBHIIOLPKLLMPENMIGHPBOHGENNJBDGOFAIADFALBPEKJAAJKAHKDLDPNEALPIPFDGOJNEHDHKMOHEBNKOCDHNIDEK',
            'cf_clearance': 'WsesLv2lPAa3euqtLmfu5Z3ioGZKZ2vVx9ARk17c8ec-1791422648-1.2.1.1-h_AsnFza4Odsn7IMnF98Q4Eh85jt70NEENBlzobntPAWKV_c8HfUp2m_VcOOy_msQ6akYkq18cIpTclJVuSdA16QaI7eNjcJ745sQrHHleogytBLlIetDFZZwnFhRKzrpCTa2DHVmIn49CwG2CczoMq1uCzZ4wrxtmbLB3iZYj0GKJoZVJkBUSbGfFokWdaHO2Rs_tNdrbZC8GR.8t5LEP146YtvTKZenl4nbJPq8bGQuUpkzqVjOIlyzWn1oiYRu8kABY.JQPacPbJrpGGv22GPo2GiSXEQQFNCvncDDn1U1HXweUGjf_.fT8sfpocKpSIVO3oWwVF_jbMcEBEG.GInfyp5SBZvV5R7J4lY20YuVeQVMgCwBRuUzMuKIpAJYZssp5IPnxaOwC4kjE9tzpPWjL33BfyF__w48HlYbg8',
            '_ga_ER466FVWPF': 'GS2.1.s1791523792$o1$g0$t1791523807$j45$l0$h0',
            '_ga': 'GA1.3.121889234.1791523793',
            '_ga_XXTTVXWHDB': 'GS2.3.s1791526012$o2$g0$t1791526012$j60$l0$h0',
            'db8ca2b43ed851cc93e71fd5fd72bff7': 'dbde1cb71be407ed7e6e7c32ea20c7f9',
            'XSRF-TOKEN': '672a1080-41d8-4bfd-86a6-99bcab9bd2c0',
            'SESSION': '65751f9c-6469-465d-a850-c902ef6dce1e',
            'f5avraaaaaaaaaaaaaaaa_session_': 'KAIDCPMFEJEFJDECMNLBCHBHAHICDCMGJIDEKOHDCJIIBLLGFLLJPCNPGIKOAEHMEPHDGHGMJPFODFGMKNGAOMEBEJIHKBOFAPJAMIHDCPBPEHDNICDLHJFMKDJMLDBD',
            'TS00000000076': '0868f8be6fab28001395d81992b92964934a13ed257e067eca6420432f67f7447bfb863cddfa73578c4ae15fdfd45ea908dca78c4209d0007679fdc96cb554a5fc406df84ee1215759242a9178a14e99a1c57ff4030174b1de3308062bea49122c6e026c71f494c017948c436ce161120c736eece1d7d366bd32409c017ed7eca8f5155137f1e46a66a5cc8ee1913a53427d99d455454c7daebf528ef63b7be51d82c8e1498865b99e17a7702043d647209d0d41e2d2c300b055f6911e103fac73b8034c9647e45081f596093b8b8aa4696ba65605093c68a18095a07354ecea1aeee6959c28e474374d082cd564bd3cf0a742f074c4beb93db530c65cda8890e182c36d96347973',
            'TSPD_101_DID': '0868f8be6fab28001395d81992b92964934a13ed257e067eca6420432f67f7447bfb863cddfa73578c4ae15fdfd45ea908dca78c42063800f03d544fda00873125ee5928f4c1575f12d35fdc6f26cae2bcbc80a31ec2c426f6324897bf8d5ede432c9ce546c7ad32f3045a232a83154b',
            'TSPD_101': '0868f8be6fab2800a255d085215043432431ef62b6561ddd5c87ad872156ec01064fe54b0910f1619ba764a8e8cc9e9208f94c8ae7051800c0f8c255225226ca5ca1732140a3428bba23ce13beb1c95e',
            'TS011f2d1a': '01266d26d08981d49e2451faf4f83c0e1b4107e43c384bea9e3b35a7ec1b11f4ee70f505030e33e0cc06c3fb0df25a20116cb48b3f',
            'TS5220f739077': '0868f8be6fab2800edac987d9ece0d67cfa82a4255e7f4fd72dd2b4e55535bd209a5fd7dcd4f746cc1fb1806b306e07a0839d82b5a172000f5171bf7587e75f1cd97bf84f94ad31b9c83c677231dbc4d08982b301beb6e9c',
            'TS5220f739029': '0868f8be6fab2800c0fb4c42e40576b2c61f61487e1504867424746ac93e6b7d6cdce20d365bde5998d6d876134d9776',
            'TSf1edb2d2027': '0868f8be6fab2000005a635acb8b3f75fcbe823e7f9ce37744b102909cc5a24e033425608cb4e74508b4715e39113000a0f97f46bba7df70115094f3fe9473a629ab02fe5b278f6f36aab61adcbd8bef296c1595ba3a79c616ef11503b75c2d0',
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
            'User-Agent': 'Mozilla/5.0 (Linux; Android 16; Pixel 10) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Mobile Safari/537.36',
            'X-XSRF-TOKEN': '672a1080-41d8-4bfd-86a6-99bcab9bd2c0',
            'sec-ch-ua': '"Google Chrome";v="155", "Chromium";v="155", "Not(A:Brand";v="24"',
            'sec-ch-ua-mobile': '?1',
            'sec-ch-ua-platform': '"Android"',
            'Cookie': 'f5avraaaaaaaaaaaaaaaa_session_=BHKPGAHHMKMBOINLLLPMLGHMDNFLPEOPMEKMBKLBHIIOLPKLLMPENMIGHPBOHGENNJBDGOFAIADFALBPEKJAAJKAHKDLDPNEALPIPFDGOJNEHDHKMOHEBNKOCDHNIDEK; cf_clearance=WsesLv2lPAa3euqtLmfu5Z3ioGZKZ2vVx9ARk17c8ec-1791422648-1.2.1.1-h_AsnFza4Odsn7IMnF98Q4Eh85jt70NEENBlzobntPAWKV_c8HfUp2m_VcOOy_msQ6akYkq18cIpTclJVuSdA16QaI7eNjcJ745sQrHHleogytBLlIetDFZZwnFhRKzrpCTa2DHVmIn49CwG2CczoMq1uCzZ4wrxtmbLB3iZYj0GKJoZVJkBUSbGfFokWdaHO2Rs_tNdrbZC8GR.8t5LEP146YtvTKZenl4nbJPq8bGQuUpkzqVjOIlyzWn1oiYRu8kABY.JQPacPbJrpGGv22GPo2GiSXEQQFNCvncDDn1U1HXweUGjf_.fT8sfpocKpSIVO3oWwVF_jbMcEBEG.GInfyp5SBZvV5R7J4lY20YuVeQVMgCwBRuUzMuKIpAJYZssp5IPnxaOwC4kjE9tzpPWjL33BfyF__w48HlYbg8; _ga_ER466FVWPF=GS2.1.s1791523792$o1$g0$t1791523807$j45$l0$h0; _ga=GA1.3.121889234.1791523793; _ga_XXTTVXWHDB=GS2.3.s1791526012$o2$g0$t1791526012$j60$l0$h0; db8ca2b43ed851cc93e71fd5fd72bff7=dbde1cb71be407ed7e6e7c32ea20c7f9; XSRF-TOKEN=672a1080-41d8-4bfd-86a6-99bcab9bd2c0; SESSION=65751f9c-6469-465d-a850-c902ef6dce1e; f5avraaaaaaaaaaaaaaaa_session_=KAIDCPMFEJEFJDECMNLBCHBHAHICDCMGJIDEKOHDCJIIBLLGFLLJPCNPGIKOAEHMEPHDGHGMJPFODFGMKNGAOMEBEJIHKBOFAPJAMIHDCPBPEHDNICDLHJFMKDJMLDBD; TS00000000076=0868f8be6fab28001395d81992b92964934a13ed257e067eca6420432f67f7447bfb863cddfa73578c4ae15fdfd45ea908dca78c4209d0007679fdc96cb554a5fc406df84ee1215759242a9178a14e99a1c57ff4030174b1de3308062bea49122c6e026c71f494c017948c436ce161120c736eece1d7d366bd32409c017ed7eca8f5155137f1e46a66a5cc8ee1913a53427d99d455454c7daebf528ef63b7be51d82c8e1498865b99e17a7702043d647209d0d41e2d2c300b055f6911e103fac73b8034c9647e45081f596093b8b8aa4696ba65605093c68a18095a07354ecea1aeee6959c28e474374d082cd564bd3cf0a742f074c4beb93db530c65cda8890e182c36d96347973; TSPD_101_DID=0868f8be6fab28001395d81992b92964934a13ed257e067eca6420432f67f7447bfb863cddfa73578c4ae15fdfd45ea908dca78c42063800f03d544fda00873125ee5928f4c1575f12d35fdc6f26cae2bcbc80a31ec2c426f6324897bf8d5ede432c9ce546c7ad32f3045a232a83154b; TSPD_101=0868f8be6fab2800a255d085215043432431ef62b6561ddd5c87ad872156ec01064fe54b0910f1619ba764a8e8cc9e9208f94c8ae7051800c0f8c255225226ca5ca1732140a3428bba23ce13beb1c95e; TS011f2d1a=01266d26d08981d49e2451faf4f83c0e1b4107e43c384bea9e3b35a7ec1b11f4ee70f505030e33e0cc06c3fb0df25a20116cb48b3f; TS5220f739077=0868f8be6fab2800edac987d9ece0d67cfa82a4255e7f4fd72dd2b4e55535bd209a5fd7dcd4f746cc1fb1806b306e07a0839d82b5a172000f5171bf7587e75f1cd97bf84f94ad31b9c83c677231dbc4d08982b301beb6e9c; TS5220f739029=0868f8be6fab2800c0fb4c42e40576b2c61f61487e1504867424746ac93e6b7d6cdce20d365bde5998d6d876134d9776; TSf1edb2d2027=0868f8be6fab2000005a635acb8b3f75fcbe823e7f9ce37744b102909cc5a24e033425608cb4e74508b4715e39113000a0f97f46bba7df70115094f3fe9473a629ab02fe5b278f6f36aab61adcbd8bef296c1595ba3a79c616ef11503b75c2d0',
        },
        "json_data": {
            'surveyPeriodId': 'e511c9a7-275d-4ce7-9776-8e65f6b4aedd',
            'surveyRoleId': '4bd73bf5-a133-4c67-9f94-fed42d72c31f',
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
            'f5avraaaaaaaaaaaaaaaa_session_': 'CLGHPLFDBOFILKLGMECDEINOIOKMGBNLFLFBFPPIJONDENDDCIADHHJOFFEINNEOCEBDKIAFMCMHLLOCCKBAMEGMAKBDIONAIMFMBLDOEOILKNCGOHGBMNHJBAMIMOPI',
            'cf_clearance': 'WsesLv2lPAa3euqtLmfu5Z3ioGZKZ2vVx9ARk17c8ec-1791422648-1.2.1.1-h_AsnFza4Odsn7IMnF98Q4Eh85jt70NEENBlzobntPAWKV_c8HfUp2m_VcOOy_msQ6akYkq18cIpTclJVuSdA16QaI7eNjcJ745sQrHHleogytBLlIetDFZZwnFhRKzrpCTa2DHVmIn49CwG2CczoMq1uCzZ4wrxtmbLB3iZYj0GKJoZVJkBUSbGfFokWdaHO2Rs_tNdrbZC8GR.8t5LEP146YtvTKZenl4nbJPq8bGQuUpkzqVjOIlyzWn1oiYRu8kABY.JQPacPbJrpGGv22GPo2GiSXEQQFNCvncDDn1U1HXweUGjf_.fT8sfpocKpSIVO3oWwVF_jbMcEBEG.GInfyp5SBZvV5R7J4lY20YuVeQVMgCwBRuUzMuKIpAJYZssp5IPnxaOwC4kjE9tzpPWjL33BfyF__w48HlYbg8',
            '_ga_ER466FVWPF': 'GS2.1.s1791523792$o1$g0$t1791523807$j45$l0$h0',
            '_ga': 'GA1.3.121889234.1791523793',
            '_ga_XXTTVXWHDB': 'GS2.3.s1791526012$o2$g0$t1791526012$j60$l0$h0',
            'db8ca2b43ed851cc93e71fd5fd72bff7': 'dbde1cb71be407ed7e6e7c32ea20c7f9',
            'XSRF-TOKEN': '672a1080-41d8-4bfd-86a6-99bcab9bd2c0',
            'SESSION': '65751f9c-6469-465d-a850-c902ef6dce1e',
            'f5avraaaaaaaaaaaaaaaa_session_': 'KAIDCPMFEJEFJDECMNLBCHBHAHICDCMGJIDEKOHDCJIIBLLGFLLJPCNPGIKOAEHMEPHDGHGMJPFODFGMKNGAOMEBEJIHKBOFAPJAMIHDCPBPEHDNICDLHJFMKDJMLDBD',
            'TS011f2d1a': '01266d26d08981d49e2451faf4f83c0e1b4107e43c384bea9e3b35a7ec1b11f4ee70f505030e33e0cc06c3fb0df25a20116cb48b3f',
            'TS5220f739078': '0868f8be6fab2000f72fd5de3aafc18643bc4840333f8a1bcbff9727f4a44ca61bc781ec33ea903a08c5b8c8da1870019dc13140922d0244f7eb7ef16c652dc4acca5b8a281cd74e81da272b401ccea4886fcfb7c9456deb1046b42926e8f340375b876e9d80dfca0282c24eb7caa249eb1442628deb2a6a618bb13c738bd29b34409b82a5428d8c93d310db1548d5491599b7598fc64026203fec8116292f048b06c3744f8ccf0952ea5a7f7f353cbac0da044453cddc6746bac14f3d687126478fec765b7ed6a933f68333ea6426f3d721914c350e78639ab8cbf8b9ba0c4d0e86b1e91f33bea91e72dfa8228534e4f17071f637cdedfd0b71b8637cec7fceb129ab9aa9e6f5f89e1f8e2e76ce135312f39f6da82da071282e7356a83fd8f6d750973d32f81e57563fb414ba399b88cc94cc79de2c20baf319233e300d7aee198bdbf47922a30e3481827788781f0c5cd49996d4a5dbe327c2d4b6feaa6936452bcdde831cbce46e2ec47f40be1a46e4dcd78bc1f034877ffbed8a6ba04ffeeb4583927aed11c23861debac3d3561b9e44604dc679f5e5183bd76b75fd8139',
            'TS00000000076': '0868f8be6fab280036035ccdcd13c9ebc9032e9ec901b886f4b20bada11273278dccd61acd523567c6df6f9b0d721c71086a35677e09d000f3c246576cfc66f66852cee00ca9cb113597fdd79a6d14bda0d4f8180edd93322f20a1a058778790a9bd205a8e9f9b321cbe5238d493d6f193c5444a0e4f02012a7f624e41aaa5e7e9067fdf9dbce824bafdc66174df6c071ece73577219611ee0ebc0af258bc3cc985596591564fc4bb006e8ff090db75d5f9ce123363e9a991ae3dc7a6dd1baae73ae6c68c4f75b691ddcb42313bd8c430be0c87e7ff5c949c3a5d03321fd65fedbc6891ef10bc95c658ca7e0ef15c6d2a56d96fda21613e9aa433b2f1b9d09b00709993f1a71bbbe',
            'TSPD_101_DID': '0868f8be6fab280036035ccdcd13c9ebc9032e9ec901b886f4b20bada11273278dccd61acd523567c6df6f9b0d721c71086a35677e063800b79ad577c5a3b320bbb84360d980df0859fe3bee34fc326e7d3ac7b405d390ccdf1650d24ba4ba8642d7340830269b4b7ec139a5f2e31f3d',
            'TSPD_101': '0868f8be6fab28005c9fa1745e98d4ff1550ffe623e34eda0fba9ea7244a580a15b879f5f827ea0967fa749fd1322ef708625100f10518009ce05cbaa0f0e42f5ca1732140a3428bba23ce13beb1c95e',
            'TS5220f739077': '0868f8be6fab28006d8cda76d8e254c385a22f78d2126114a7da11ec4c670cd1a554a90ae1af93670cf94dcdddc06230086cbd36d917200082e26bc7cee44c25ff73e76b368f4e9363de747c5ce4243b0077022a2f76aab1',
            'TS5220f739029': '0868f8be6fab2800b4f6a8465a8608b81232925fb1384719f7041a489c8619dc0088e83dddaa88346d6bc2269824c843',
            'TSf1edb2d2027': '0868f8be6fab2000e719bdf571944523156064c291b7fa795e3d044352844cd6cd23f54d4693d8f508518d1278113000aaa20ca1bbccf7f3b3a667d8aa03c572e61b68eb62c3a402640307d3e764ff5dc94294a199dfe60ab6e4fc3f3edb4288',
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
            'User-Agent': 'Mozilla/5.0 (Linux; Android 16; Pixel 10) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Mobile Safari/537.36',
            'X-XSRF-TOKEN': '672a1080-41d8-4bfd-86a6-99bcab9bd2c0',
            'sec-ch-ua': '"Google Chrome";v="155", "Chromium";v="155", "Not(A:Brand";v="24"',
            'sec-ch-ua-mobile': '?1',
            'sec-ch-ua-platform': '"Android"',
            'Cookie': 'f5avraaaaaaaaaaaaaaaa_session_=CLGHPLFDBOFILKLGMECDEINOIOKMGBNLFLFBFPPIJONDENDDCIADHHJOFFEINNEOCEBDKIAFMCMHLLOCCKBAMEGMAKBDIONAIMFMBLDOEOILKNCGOHGBMNHJBAMIMOPI; cf_clearance=WsesLv2lPAa3euqtLmfu5Z3ioGZKZ2vVx9ARk17c8ec-1791422648-1.2.1.1-h_AsnFza4Odsn7IMnF98Q4Eh85jt70NEENBlzobntPAWKV_c8HfUp2m_VcOOy_msQ6akYkq18cIpTclJVuSdA16QaI7eNjcJ745sQrHHleogytBLlIetDFZZwnFhRKzrpCTa2DHVmIn49CwG2CczoMq1uCzZ4wrxtmbLB3iZYj0GKJoZVJkBUSbGfFokWdaHO2Rs_tNdrbZC8GR.8t5LEP146YtvTKZenl4nbJPq8bGQuUpkzqVjOIlyzWn1oiYRu8kABY.JQPacPbJrpGGv22GPo2GiSXEQQFNCvncDDn1U1HXweUGjf_.fT8sfpocKpSIVO3oWwVF_jbMcEBEG.GInfyp5SBZvV5R7J4lY20YuVeQVMgCwBRuUzMuKIpAJYZssp5IPnxaOwC4kjE9tzpPWjL33BfyF__w48HlYbg8; _ga_ER466FVWPF=GS2.1.s1791523792$o1$g0$t1791523807$j45$l0$h0; _ga=GA1.3.121889234.1791523793; _ga_XXTTVXWHDB=GS2.3.s1791526012$o2$g0$t1791526012$j60$l0$h0; db8ca2b43ed851cc93e71fd5fd72bff7=dbde1cb71be407ed7e6e7c32ea20c7f9; XSRF-TOKEN=672a1080-41d8-4bfd-86a6-99bcab9bd2c0; SESSION=65751f9c-6469-465d-a850-c902ef6dce1e; f5avraaaaaaaaaaaaaaaa_session_=KAIDCPMFEJEFJDECMNLBCHBHAHICDCMGJIDEKOHDCJIIBLLGFLLJPCNPGIKOAEHMEPHDGHGMJPFODFGMKNGAOMEBEJIHKBOFAPJAMIHDCPBPEHDNICDLHJFMKDJMLDBD; TS011f2d1a=01266d26d08981d49e2451faf4f83c0e1b4107e43c384bea9e3b35a7ec1b11f4ee70f505030e33e0cc06c3fb0df25a20116cb48b3f; TS5220f739078=0868f8be6fab2000f72fd5de3aafc18643bc4840333f8a1bcbff9727f4a44ca61bc781ec33ea903a08c5b8c8da1870019dc13140922d0244f7eb7ef16c652dc4acca5b8a281cd74e81da272b401ccea4886fcfb7c9456deb1046b42926e8f340375b876e9d80dfca0282c24eb7caa249eb1442628deb2a6a618bb13c738bd29b34409b82a5428d8c93d310db1548d5491599b7598fc64026203fec8116292f048b06c3744f8ccf0952ea5a7f7f353cbac0da044453cddc6746bac14f3d687126478fec765b7ed6a933f68333ea6426f3d721914c350e78639ab8cbf8b9ba0c4d0e86b1e91f33bea91e72dfa8228534e4f17071f637cdedfd0b71b8637cec7fceb129ab9aa9e6f5f89e1f8e2e76ce135312f39f6da82da071282e7356a83fd8f6d750973d32f81e57563fb414ba399b88cc94cc79de2c20baf319233e300d7aee198bdbf47922a30e3481827788781f0c5cd49996d4a5dbe327c2d4b6feaa6936452bcdde831cbce46e2ec47f40be1a46e4dcd78bc1f034877ffbed8a6ba04ffeeb4583927aed11c23861debac3d3561b9e44604dc679f5e5183bd76b75fd8139; TS00000000076=0868f8be6fab280036035ccdcd13c9ebc9032e9ec901b886f4b20bada11273278dccd61acd523567c6df6f9b0d721c71086a35677e09d000f3c246576cfc66f66852cee00ca9cb113597fdd79a6d14bda0d4f8180edd93322f20a1a058778790a9bd205a8e9f9b321cbe5238d493d6f193c5444a0e4f02012a7f624e41aaa5e7e9067fdf9dbce824bafdc66174df6c071ece73577219611ee0ebc0af258bc3cc985596591564fc4bb006e8ff090db75d5f9ce123363e9a991ae3dc7a6dd1baae73ae6c68c4f75b691ddcb42313bd8c430be0c87e7ff5c949c3a5d03321fd65fedbc6891ef10bc95c658ca7e0ef15c6d2a56d96fda21613e9aa433b2f1b9d09b00709993f1a71bbbe; TSPD_101_DID=0868f8be6fab280036035ccdcd13c9ebc9032e9ec901b886f4b20bada11273278dccd61acd523567c6df6f9b0d721c71086a35677e063800b79ad577c5a3b320bbb84360d980df0859fe3bee34fc326e7d3ac7b405d390ccdf1650d24ba4ba8642d7340830269b4b7ec139a5f2e31f3d; TSPD_101=0868f8be6fab28005c9fa1745e98d4ff1550ffe623e34eda0fba9ea7244a580a15b879f5f827ea0967fa749fd1322ef708625100f10518009ce05cbaa0f0e42f5ca1732140a3428bba23ce13beb1c95e; TS5220f739077=0868f8be6fab28006d8cda76d8e254c385a22f78d2126114a7da11ec4c670cd1a554a90ae1af93670cf94dcdddc06230086cbd36d917200082e26bc7cee44c25ff73e76b368f4e9363de747c5ce4243b0077022a2f76aab1; TS5220f739029=0868f8be6fab2800b4f6a8465a8608b81232925fb1384719f7041a489c8619dc0088e83dddaa88346d6bc2269824c843; TSf1edb2d2027=0868f8be6fab2000e719bdf571944523156064c291b7fa795e3d044352844cd6cd23f54d4693d8f508518d1278113000aaa20ca1bbccf7f3b3a667d8aa03c572e61b68eb62c3a402640307d3e764ff5dc94294a199dfe60ab6e4fc3f3edb4288',
        },
        "json_data": {
            'surveyPeriodId': '860befa8-a519-4571-b79d-4fd745d9f0a8',
            'surveyRoleId': '28a6c674-9781-48b7-92fb-c612393a7b88',
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