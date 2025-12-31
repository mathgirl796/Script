function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
async function run_wait(func, ms) {
    func()
    await sleep(ms)
}

for(i=0;i<3;i++){
    await run_wait(()=>{
    console.log(i)
    },1000)
    await run_wait(()=>{
    console.log(-i)
    },1000)
}
