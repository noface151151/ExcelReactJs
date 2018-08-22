const WriteLog=(content)=>{
    if(process.env.NODE_ENV === "development"){
        console.log(content);
    }
    return;
}
export default WriteLog;