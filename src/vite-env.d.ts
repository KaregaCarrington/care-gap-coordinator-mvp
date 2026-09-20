declare global { interface Window { api: { loadRules(): Promise<string>; saveRules(c:string):Promise<void>; saveDialog(n:string):Promise<{canceled:boolean,filePath?:string}>; writeFile(p:string,b:Uint8Array):Promise<boolean>; version():Promise<string> } } }
export {};
