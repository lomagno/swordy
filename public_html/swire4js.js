/* global msgpack */

function swire() {
}

/**
 * 
 * @param {json} jsonRequest
 * @returns {void}
 */
swire.encode = function(jsonRequest) {
    var u8Request = msgpack.encode(jsonRequest);
    return btoa(String.fromCharCode.apply(null, u8Request));        
};

/*
 * 
 * @param {string} base64String
 * @returns {void}
 */
swire.decode = function(base64String) {
    var u8 = new Uint8Array(atob(base64String).split("").map(function(c) {
        return c.charCodeAt(0);
    }));
    return msgpack.decode(u8);        
};

/**
 * 
 * @param {string} value
 * @returns {number}
 */
swire.getMissingValue = function(value) {
    switch (value) {
        case undefined:
            return 8.98846567431158e+307;
        case 'a':
            return 8.990660123939097e+307;
        case 'b':
            return 8.992854573566614e+307;
        case 'c':
            return 8.995049023194132e+307;
        case 'd':
            return 8.99724347282165e+307;
        case 'e':
            return 8.999437922449167e+307;
        case 'f':
            return 9.001632372076684e+307;
        case 'g':
            return 9.003826821704202e+307;
        case 'h':
            return 9.00602127133172e+307;
        case 'i':
            return 9.008215720959237e+307;
        case 'j':
            return 9.010410170586754e+307;
        case 'k':
            return 9.012604620214272e+307;
        case 'l':
            return 9.01479906984179e+307;
        case 'm':
            return 9.016993519469307e+307;
        case 'n':
            return 9.019187969096824e+307;
        case 'o':
            return 9.021382418724342e+307;
        case 'p':
            return 9.02357686835186e+307;
        case 'q':
            return 9.025771317979377e+307;
        case 'r':
            return 9.027965767606894e+307;
        case 's':
            return 9.030160217234412e+307;
        case 't':
            return 9.03235466686193e+307;
        case 'u':
            return 9.034549116489447e+307;
        case 'v':
            return 9.036743566116964e+307;
        case 'w':
            return 9.038938015744481e+307;
        case 'x':
            return 9.041132465371999e+307;
        case 'y':
            return 9.043326914999516e+307;
        case 'z':
            return 9.045521364627034e+307;
        default:
            return 8.98846567431158e+307;
    }
};

