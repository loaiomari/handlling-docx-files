const express = require("express")

const cors = require("cors")

const fileRoutes=require("./routes/fileRoutes")


const app = express()

app.use(cors())

app.use(express.json())

app.use(express.static("public")) // this is for front end 

app.use("/api",fileRoutes)

app.listen(3000,()=>{
    console.log("server has been started at port 3000")
})