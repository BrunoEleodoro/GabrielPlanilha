require('dotenv').config({ path: 'slack' });

var Botkit = require('botkit');
var express = require('express');
const https = require('https');
var fs = require('fs')
var path = require('path')
const { exec } = require('child_process');

// Configure your bot.
var slackController = Botkit.slackbot({ clientSigningSecret: process.env.SLACK_SIGNING_SECRET });
var slackBot = slackController.spawn({
    token: process.env.SLACK_TOKEN
});
// slackController.hears(['.*'], ['direct_message', 'direct_mention', 'mention'], function(bot, message) {
slackController.hears(['.*'], ['direct_message', 'direct_mention', 'other_event', 'file_shared'], function (bot, message) {
    slackController.log('Slack message received');
    console.log('message', message);
    // bot.replyInThread(message, 'hahaha')
    bot.api.files.upload({
        file: fs.createReadStream("main.py"),
        filename: "file1" + ".py",
        filetype: "python",
        channels: message.channel
    }, function (err, res) {
        if (err) {
            console.log("Failed to add file :(", err)
            bot.reply(message, 'Sorry, there has been an error: ' + err)
        }
    })
    exec('chmod -R 777 criar_planilha.js', (err, stdout, stderr) => {console.log(stderr)});
    exec('chmod -R 777 bot.js', (err, stdout, stderr) => {console.log(stderr)});
    exec('chmod -R 777 *.*', (err, stdout, stderr) => {console.log(stderr)});
    // bot.send("Shutting down VM #34324....", "fdsafdsafdsa");
    // bot.send("Shutting down VM #34324....")
    // var slackMessage = 
    // middleware.interpret(bot, message, function() {
    //   if (message.watsonError) {
    //     console.log(message.watsonError);
    //     bot.reply(message, message.watsonError.description || message.watsonError.error);
    //   } else if (message.watsonData && 'output' in message.watsonData) {
    //     bot.reply(message, "fdsafdsafdsa34243kj2n4jkn43jkndjksnfdsakjnfdjksanfkjdsn");
    //   } else {
    //     console.log('Error: received message in unknown format. (Is your connection with Watson Assistant up and running?)');
    //     bot.reply(message, 'I\'m sorry, but for technical reasons I can\'t respond to your message');
    //   }
});
slackController.on('file_shared', function (bot, message) {

    bot.api.files.info({ file: message.file_id }, (err, response) => {
        // console.log(response.file.title)
        console.log(response.file.filetype)
        if (response.file.title == "config_criar_planilha") {
            const file = fs.createWriteStream(path.join(__dirname,"config_criar_planilha"));
            console.log(path.join(__dirname,"config_criar_planilha"))
            https.get(response.file.url_private_download, {
                headers: {
                    'Authorization': 'Bearer ' + process.env.SLACK_TOKEN
                }
            }, function (response) {
                response.pipe(file);
                console.log(fs.existsSync(path.join(__dirname,"config_criar_planilha")))
            });
            bot.say({
                text: "Config file for 'Criar planilha' Received! :fbhappy:",
                channel: message.channel_id // channel Id for #slack_integration
            });
        } else if (response.file.title == "config") {
            const file = fs.createWriteStream(path.join(__dirname,"config"));
            https.get(response.file.url_private_download, {
                headers: {
                    'Authorization': 'Bearer ' + process.env.SLACK_TOKEN
                }
            }, function (response) {
                response.pipe(file);
            });
            bot.say({
                text: "Config file for all the scripts Received! :fbhappy:",
                channel: message.channel_id // channel Id for #slack_integration
            });
        } else if (response.file.filetype == "csv") {
            const file = fs.createWriteStream(path.join(__dirname,"a.csv"))
            https.get(response.file.url_private_download, {
                headers: {
                    'Authorization': 'Bearer ' + process.env.SLACK_TOKEN
                }
            }, function (response) {
                response.pipe(file);
                build(bot, message);
            });
            bot.say({
                text: "Received! :fbhappy: \nWorking now... :construction-2:",
                channel: message.channel_id // channel Id for #slack_integration
            });
        }

    })
});

slackBot.startRTM();

function build(bot, message) {
    // exec('node criar_planilha.js', (err, stdout, stderr) => {
    //     console.log('aaa');
    //     console.log(stderr);
    //     console.log(stdout);
    //     exec('ls -la', (err, stdout, stderr) => {
    //         console.log('lista de arquivos')
    //         console.log(stdout)
    //     })

    // })
    // console.log(message.channel)
    // cmd.exe /c executar_todos.bat
    exec('make build', (err, stdout, stderr) => {
        console.log(stdout)
        fs.readFile("config", { encoding: 'utf-8' }, function (err, data) {
            var file_name = data.split('\n')[2].replace("OUTPUT_FILE=", "")
            if(fs.existsSync(path.join(__dirname, file_name))) {
	            bot.api.files.upload({
	                file: fs.createReadStream(path.join(__dirname, file_name)),
	                filename: file_name,
	                filetype: "xlsx",
	                channels: message.channel_id
	            }, function (err, res) {
	                if (err) {
	                    console.log("Failed to add file :(", err)
	                    bot.reply(message, 'Sorry, there has been an error: ' + err)
	                }
	            })	
            } else {
            	console.log('file dont exists',path.join(__dirname, file_name))
            }
        });


    });

}

// // Create an Express app
// var app = express();
// var port = process.env.PORT || 5000;
// app.set('port', port);
// app.listen(port, function () {
//     console.log('Client server listening on port ' + port);
// });
