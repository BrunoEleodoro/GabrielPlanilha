require("./bot.js")
require("./bot_grafana.js")
var express = require('express');

// // Create an Express app
var app = express();
var port = process.env.PORT || 5000;
app.set('port', port);
app.get('/', (req, res) => {
    res.send('working')
})
app.listen(port, function () {
    console.log('Client server listening on port ' + port);
});
