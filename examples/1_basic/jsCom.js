var wsh = new ActiveXObject("WScript.Shell");

// Add a JS helper that uses the actual COM method
wsh.safeExpandEnv = function(varName) {
    var result = this.ExpandEnvironmentStrings("%" + varName + "%");
    if (result === "%" + varName + "%") {
        return "(undefined)";  // Variable doesn't exist
    }
    return result;
};

alert(wsh)
alert(wsh.safeExpandEnv)

// Now you can call it!
console.log(wsh.safeExpandEnv("PATH"));  // Returns expanded path or "(undefined)"
