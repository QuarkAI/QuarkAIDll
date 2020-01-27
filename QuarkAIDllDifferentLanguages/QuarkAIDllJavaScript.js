var MyQuarkAI = new ActiveXObject("QuarkAI.dll");
var MyResult;
MyResult =  MyQuarkAI.Message("This is an example of using QuarkAI.dll in JavaScript.","QuarkAIJaveScript", 0)
delete MyQuarkAI;