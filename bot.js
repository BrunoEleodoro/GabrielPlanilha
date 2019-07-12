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
    bot.reply(message, "I'm here :) :hello-bear:");
});
slackController.on('file_shared', function (bot, message) {

    bot.api.files.info({ file: message.file_id }, (err, response) => {
        // console.log(response.file.title)
        console.log(response.file.filetype)
        if (response.file.title == "config_criar_planilha") {
            const file = fs.createWriteStream(path.join(__dirname, "config_criar_planilha"));
            console.log(path.join(__dirname, "config_criar_planilha"))
            https.get(response.file.url_private_download, {
                headers: {
                    'Authorization': 'Bearer ' + process.env.SLACK_TOKEN
                }
            }, function (response) {
                response.pipe(file);
                console.log(fs.existsSync(path.join(__dirname, "config_criar_planilha")))
            });
            bot.say({
                text: "Config file for 'Criar planilha' Received! :fbhappy:",
                channel: message.channel_id // channel Id for #slack_integration
            });
        } else if (response.file.title == "config") {
            const file = fs.createWriteStream(path.join(__dirname, "config"));
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
            if (response.file.title.includes("metrics_")) {
                var output_filename = response.file.title;
                output_filename = output_filename.split(".")[0];
                output_filename = output_filename.split("metrics_");
                output_filename = output_filename[1].toString().trim().split("-")
                output_filename = output_filename[2] + "-" + output_filename[1] + "-" + output_filename[0]
                output_filename = "Metricas_" + output_filename + ".xlsx"
                const file = fs.createWriteStream(path.join(__dirname, "a.csv"))
                https.get(response.file.url_private_download, {
                    headers: {
                        'Authorization': 'Bearer ' + process.env.SLACK_TOKEN
                    }
                }, function (response) {
                    response.pipe(file);
                    build(bot, message, output_filename);
                });
                bot.say({
                    text: "Received! :fbhappy: \nProcessing file and collecting metrics... :construction-2:",
                    channel: message.channel_id // channel Id for #slack_integration
                });
            } else {
                bot.say({
                    text: "Invalid file name :sad1: ",
                    channel: message.channel_id // channel Id for #slack_integration
                });
            }

        }

    })
});

slackBot.startRTM();

function build(bot, message, output_filename) {
    exec('make build', (err, stdout, stderr) => {
        console.log(stdout)
        fs.readFile("config", { encoding: 'utf-8' }, function (err, data) {
            var file_name = data.split('\n')[2].replace("OUTPUT_FILE=", "")
            if (fs.existsSync(path.join(__dirname, file_name))) {
                bot.say({
                    text: "Finished, uploading the file...",
                    channel: message.channel_id // channel Id for #slack_integration
                });
                bot.api.files.upload({
                    file: fs.createReadStream(path.join(__dirname, file_name)),
                    filename: output_filename,
                    filetype: "xlsx",
                    channels: message.channel_id,

                }, function (err, res) {
                    if (err) {
                        console.log("Failed to add file :(", err)
                        bot.reply(message, 'Sorry, there has been an error: ' + err)
                    }
                })
            } else {
                console.log('file dont exists', path.join(__dirname, file_name))
            }
        });


    });

}

var app = express();
var port = process.env.PORT || 5000;
app.set('port', port);
app.get('/', (req, res) => {
    res.send('working')
})
app.listen(port, function () {
    console.log('Client server listening on port ' + port);
});
