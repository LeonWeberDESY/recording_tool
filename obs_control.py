import sys
import asyncio
import simpleobsws
import logging
import os
import json

# Open and read the JSON file
with open('config.json', 'r') as json_file:
    cfg = json.loads(json_file.read())
    host = cfg["host"]
    port = cfg["port"]
    password = cfg["password"]
    scene_name = cfg["scene_name"]
    input_name = cfg["input_name"]
    device_id = cfg["device_id"]




LOGFILE_PATH = os.path.join(os.path.dirname(__file__), "logs_obs_control.log")


# Configure logging with timestamps
logging.basicConfig(
    filename=LOGFILE_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)


async def main():
    if len(sys.argv) < 2:
        logging.warning("Invalid action, use 'start' or 'stop' as arguments")
        return
    
    action = sys.argv[1].lower()
    ws = simpleobsws.WebSocketClient(
        url=f"ws://{host}:{port}",
        password=password
    )
    await ws.connect()
    await ws.wait_until_identified()

    try:
        if action == "start":
            logging.info(f"Creating mic input '{input_name}' in scene '{scene_name}'...")
            await ws.call(simpleobsws.Request("CreateInput", {
                "sceneName": scene_name,
                "inputName": input_name,
                "inputKind": "wasapi_input_capture",
                "inputSettings": {"device_id": device_id},
                "sceneItemEnabled": True
            }))
            await ws.call(simpleobsws.Request("StartRecord"))
            logging.info("Recording started")

        elif action == "stop":
            await ws.call(simpleobsws.Request("StopRecord"))
            await ws.call(simpleobsws.Request("RemoveInput", {"inputName": input_name}))
            logging.info("Recording stopped and mic input removed")

        else:
            logging.error("Invalid action, use 'start' or 'stop' as arguments")

    finally:
        await ws.disconnect()

if __name__ == "__main__":
    asyncio.run(main())
