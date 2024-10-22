# -*- coding: utf-8 -*-
import json
import random
import time
import uuid
import requests
import websocket
from threading import Thread
from base64 import b64decode, b64encode
from requests.adapters import HTTPAdapter
from datetime import datetime
from urllib.parse import urlencode, urljoin


class WebSocketClient:
    def __init__(self, url):
        self.wss_url = url
        self.wss = None
        self.keep_running = True
        self.last_ping_time = None
        self.lock = False

    def keeplive(self, wss):
        while self.keep_running:
            time.sleep(15)
            current_time = int(time.time() * 1000)
            if self.last_ping_time is None or current_time - self.last_ping_time >= 15000:  # 每15秒发送一次ping
                self.send_message(r'{"type":6}'+chr(30))
                self.last_ping_time = current_time

    def on_open(self, wss):
        print("Connection opened")
        Thread(target=self.keeplive, args=(wss,)).start()

    def on_message(self, wss, message):
        data = json.loads(message)
        if data.get('type') == 1:
            message = data["arguments"]["messages"]["text"]  # 输出内容
            # data["arguments"]["messages"]["adaptiveCards"][0]["body"][0]["text"]  # 输出内容 需要增加格式判断
            print(message, end="", flush=True)
            self.lock = True
            return
        if data.get('type') == 2:
            print("\n输出完毕")
            self.lock = False
            return
        print(f"Received message: {message}")
        self.lock = False
        return

    def send_message(self, message):
        while self.keep_running:
            if self.lock is False:
                self.lock = True
                break
            else:
                time.sleep(0.5)

        if self.wss and self.wss.sock and self.wss.sock.connected:
            self.wss.send(message)
        else:
            print("WebSocket is not connected.")
            self.lock = False

    def on_error(self, wss, error):
        print(f"Error: {error}")

    def on_close(self, wss, close_status_code, close_msg):
        print("Connection closed")
        self.keep_running = False

    def on_pong(self, wss, message):
        print(f"Received Pong message: {message}")

    def on_ping(self, wss, message):
        print(f"Received Pong message: {message}")

    def run(self):
        while self.keep_running:
            try:
                headers = {
                    "Host": "substrate.office.com",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
                    "Accept": "*/*",
                    "Accept-Language": "zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
                    "Accept-Encoding": "gzip, deflate, br",
                    "Sec-WebSocket-Version": "13",
                    "Origin": "https://outlook.office.com",
                    "Sec-WebSocket-Extensions": "permessage-deflate",
                    # Sec-WebSocket-Key: Cj1Lsw==
                    "Connection": "keep-alive, Upgrade",
                    "Cookie": "ClientId=ccc; OIDC=1; MUID=cccc",
                    "Pragma": "no-cache",
                    "Cache-Control": "no-cache",
                    "Upgrade": "websocket",
                }
                self.wss = websocket.WebSocketApp(self.wss_url,
                                                  # header=headers,
                                                  on_open=self.on_open,
                                                  on_message=self.on_message,
                                                  on_error=self.on_error,
                                                  on_close=self.on_close)
                # self.wss.on_pong = self.on_pong
                # self.wss.on_ping = self.on_ping
                # self.wss.run_forever(ping_interval=5, ping_timeout=3)
                self.wss.run_forever()
            except Exception as e:
                print(f"Exception: {e}")
                time.sleep(5)  # Wait before reconnecting

    def start(self):
        self.keep_running = True
        t = Thread(target=self.run)
        t.setDaemon(True)
        t.start()

    def stop(self):
        self.keep_running = False
        if self.wss:
            self.wss.close()


class Copilot(WebSocketClient):
    def __init__(self, url=None):
        super().__init__(url)
        self.invocationId = 2
        self.isStartOfSession = False

    def get_wss_url(self, X_SessionId, ClientRequestId, access_token, ConversationId=None):
        wss = "wss://substrate.office.com"
        path = "Oid:xxx9725-xxx"
        data = {
            "scenario": "OfficeWebFreeCopilot",
            "variants": "feature.cwcallowedos",
            "X-SessionId": X_SessionId,
            "ConversationId": ConversationId,
            "ClientRequestId": ClientRequestId,
            "source": "officeweb",
            "access_token": access_token
        }
        if not data["ConversationId"]:
            data.pop("ConversationId")
        params = urlencode(data)
        wss_url = f"{wss}/m365Copilot/ChatHub/{path}?{params}"
        return wss_url

    def prepMessageToSend(self, message, X_SessionId, ClientRequestId):
        _1 = {"protocol": "json", "version": 1}
        self.send_message(_1)

        _2 = {"type": 6}
        self.send_message(_2)

        if self.isStartOfSession == True:
            self.invocationId = 0
        message_data = {
            "arguments": [
                {
                    "optionsSets": ["cwc_flux_image", "flux_fileupload_odb", "ldsummary", "ldqa", "sdretrieval"],
                    "allowedMessageTypes": ["ActionRequest", "Chat", "ConfirmationCard", "Context",
                                            "InternalSearchQuery",
                                            "InternalSearchResult", "Disengaged", "InternalLoaderMessage", "Progress",
                                            "RenderCardRequest", "RenderContentRequest", "AdsQuery", "SemanticSerp",
                                            "GenerateContentQuery", "SearchQuery", "GeneratedCode",
                                            "InternalTasksMessage",
                                            "Disclaimer", "RecommendPlugin"],
                    "sliceIds": [],
                    "plugins": [{
                        "Id": "BingWebSearch",
                        "Source": "BuiltIn"
                    }],
                    "threadType": "webchat",
                    "sessionId": X_SessionId,
                    "source": "officeweb",
                    "scenario": "OfficeWebFreeCopilot",
                    "clientInfo": {
                        "clientPlatform": "mcmcopilot-web",
                        "clientAppName": "Office",
                        "clientEntrypoint": "mcmcopilot-officeweb",
                        "clientSessionId": X_SessionId
                    },
                    "isStartOfSession": self.isStartOfSession,
                    "traceId": ClientRequestId,
                    "requestId": ClientRequestId,
                    "message": {
                        "locale": "en-US",
                        "market": "en-US",
                        "region": "US",
                        "location": "lm;",
                        "adaptiveCards": [],
                        "author": "user",
                        "inputMethod": "Keyboard",
                        "text": message,
                        "messageType": "Chat",
                        "requestId": ClientRequestId,
                        "messageId": ClientRequestId,
                        "locationInfo": {
                            "timeZoneOffset": 8
                        }
                    },
                    "tone": "Creative",
                    "spokenTextMode": "None"
                }],
            "invocationId": str(self.invocationId),
            "target": "chat",
            "type": 4
        }
        self.send_message(message_data)

    def preSend_message_time(self):
        unow = time.time()  # 1
        cnow = unow + random.uniform(0.015, 0.025)  # 2
        c2now = unow + random.uniform(0.8, 1.5)  # 3
        rnow = c2now + random.uniform(0.001, 0.003)  # 4
        timestamps = {
            "ConnectionStart": datetime.utcfromtimestamp(cnow).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z',  # 2
            "ConnectionEstablished": datetime.utcfromtimestamp(c2now).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z',  # 3
            "UserInputSubmit": datetime.utcfromtimestamp(unow).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z',  # 1
            "RequestSent": datetime.utcfromtimestamp(rnow).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z',  # 4
        }
        data = {
            "arguments":
            [{"timestamps": timestamps}],
            "target": "Metrics",
            "type": 1
        }
        self.send_message(data)
        return timestamps

    def run_task(self):
        # 流程 "protocol":"json" 1-1> {"type":6}-arguments-arguments 3-n> arguments-{"type":7} 2-0> over
        X_SessionId = "359a0779a85f"
        ConversationId = "b5e04413f50"
        # ClientRequestId = "ccc-e-ccc"    # 每次问答都变 token也是
        ClientRequestId = uuid.uuid4().__str__()    # 每次问答都变 token也是
        access_token = self.get_token()["access_token"]
        while True:
            message = input("请输入内容：")
            if message == "q":
                self.stop()
                break
            self.wss_url = self.get_wss_url(X_SessionId, ClientRequestId, access_token, ConversationId)
            print(self.wss_url)
            self.start()

            self.prepMessageToSend(message, X_SessionId, ClientRequestId)

            timestamps = self.preSend_message_time()

            time.sleep(1)   # 等待响应输出完成

            self.afterMessageReceived(timestamps)
            self.stop()

    def afterMessageReceived(self, timestamps):
        c2now = time.time()
        fnow = c2now + random.uniform(1.5, 2.3)  # 1
        lnow = fnow + random.uniform(1, 1.5)  # 2
        l2now = lnow + random.uniform(0.5, 1.2)  # 3
        timestamps.update({
            "FirstTokenReceived": datetime.utcfromtimestamp(fnow).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z',  # 1
            "FirstTokenRendered": datetime.utcfromtimestamp(lnow).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z',  # 2
            "LastTokenReceived": datetime.utcfromtimestamp(l2now).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z',  # 3
            "LastTokenRendered": datetime.utcfromtimestamp(lnow).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + 'Z',  # 2
        })
        data = {
            "arguments":
            [{"timestamps": timestamps}],
            "target": "Metrics",
            "type": 1
        }
        self.send_message(data)

        _x = {"type": 7}
        self.send_message(_x)

        self.invocationId += 1  # 计数器

    def get_token(self):
        url = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"
        url = "https://login.microsoftonline.com/d0c6deccc/oauth2/v2.0/token"
        # client-request-id 不同，变化规则？
        data = {
            "client_id": "xx-xxx-xx-xx-xxx",
            "scope": "https://substrate.office.com/sydney/.default openid profile offline_access",
            "grant_type": "refresh_token",
            "client_info": "1",
            "x-client-SKU": "msal.js.browser",
            "x-client-VER": "3.22.0",
            "x-ms-lib-capability": "retry-after, h429",
            "x-client-current-telemetry": "5|,|,",
            "x-client-last-telemetry": "5|2|0",
            # "client-request-id": "910412xxxbf53d",
            "client-request-id": uuid.uuid4().__str__(),
            "refresh_token": "0.xxx.",
            "X-AnchorMailbox": "Oid:4a9699725-cc"
        }
        headers = {
            "Host": "login.microsoftonline.com",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
            "Accept": "*/*",
            "Accept-Language": "zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
            "Accept-Encoding": "gzip, deflate, br",
            "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
            "Origin": "https://outlook.office.com",
        }
        data = urlencode(data)
        print(data)
        with requests.post(url=url, headers=headers, data=data) as resp:
            print(resp.status_code)
            print(resp.headers)
            print(resp.text)
            return resp.json()
        # {"token_type":"Bearer","scope":"https://substrate.office.com/sydney/M365Chat.Read https://substrate.office.com/sydney/sydney.readwrite https://substrate.office.com/sydney/.default","expires_in":5205,"ext_expires_in":5205,"access_token":"..d4Zf2KECnbEdUE---IGmbf"}

def main():
    copilot = Copilot()
    copilot.run_task()


# 使用示例
if __name__ == "__main__":
    main()
    exit()
    # {"token_type":"Bearer","scope":"https://substra

    url = "https://login.microsoftonline.com/d9a/oauth2/v2.0/token"
    data = "cl"
    # {"token_type":"Bearer","scope":"https://www.office.com/v2/OfficeHome.All","expires_in":3871,"ext_expires_in":3871,"access_token":"e"}

    # https://substrate.office.com/sydney/.default openid profile offline_access
    # 27997/.default openid profile offline_access
    # https://www.office.com/v2/OfficeHome.All openid profile offline_access

    tnow = time.time()

    x = datetime.utcfromtimestamp(tnow)
    y = datetime.utcfromtimestamp(tnow + random.uniform(0.015, 0.025))
    print(x)
    print(y)

    # 第一次请求
# wss://substrate.office.com/m365Copilot/ChatHub/O
# 发 {"protocol":"json","version":1}
# {}
# 发 {"type":6} # ping  15s间隔
# 发 # 发 # 发 # 发 {"type":7} # 关闭
# 结束

# 第二次
# wss://substrate.office.com/m365Copilot/ChatHub/
# 发 {"protocol":"json","version":1}
# {}
# 发 {"type":6}
# 发# 发# 发# 发 {"type":7}
# 结束


# 检测token有效性
# https://m365.cloud.microsoft/api/checklogin/v3?hybridauth=1&workload=officehomereact&auth=2
# hybridauth=1&workload=officehomereact&auth=2
# {"LoginState":"SignedIn"}


# 刷新token1
# https://login.microsoftonline.com/d0c6
# client_id=4
# {"token_typ

# 刷新token2  55分钟刷新一次
# https://login.microsoftonline.com/d0c9a/oauth2/v2.0/token
# client_i
# {"token_

#########################
# https://login.microsoftonline.com/d0c6de7f/oauth2/v2.0/token
# client_id=47
# {"token_type":"

# https://login.microsoftonline.com/organizations/oauth2/v2.0/token
# client_id
# {"token_t

# https://m365.cloud.microsoft/api/checklogin/v3?hybridauth=1&workload=officehomereact&auth=2
# GET /api/checklogin/v3?hybridauth=1&workload=officehomereact&auth=2 HTTP/2
# Host: m365.cloud.microsoft
# User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0
# Accept: application/json, text/plain, */*
# Accept-Language: zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2
# Accept-Encoding: gzip, deflate, br
# Referer: https://m365.cloud.microsoft/chat/?auth=2&home=1&from=NoAccountOnToken
# X-OfficeHome-UserId: 
# X-OfficeHome-TenantId: d0c6
# X-OfficeHome-AuthVersion: 2.0
# X-OfficeHome-CorrelationId: b1
# Connection: keep-alive
# Cookie: OH.FLID=a1
# TE: trailers
# {"LoginState":"SignedIn"}





