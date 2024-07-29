from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
import uvicorn
from md2ppt.markdown2json import router as markdown2json_router

app = FastAPI(title="Markdown to PowerPoint API", version="1.0.0")


app.include_router(router=markdown2json_router,prefix="/md2ppt")
app.mount("/pptSource", StaticFiles(directory="pptSource"), name="pptSource")


@app.get("/")
def read_root():
    return {"msg": "Hello World"}



if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)


