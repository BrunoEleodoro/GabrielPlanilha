require('dotenv').config({ path: 'config_grafana' });

var Botkit = require('botkit');
var express = require('express');
const https = require('https');
var fs = require('fs')
var path = require('path')
const { exec } = require('child_process');

// Configure your bot.
var slackController = Botkit.slackbot({ clientSigningSecret: process.env.GRAFANA_SLACK_SIGNING_SECRET });
var slackBot = slackController.spawn({
    token: process.env.GRAFANA_SLACK_TOKEN
});
// slackController.hears(['.*'], ['direct_message', 'direct_mention', 'mention'], function(bot, message) {
slackController.hears(['.*'], ['direct_message', 'direct_mention', 'other_event', 'file_shared'], function (bot, message) {
    slackController.log('Slack message received');
    // console.log('message', message);
    // bot.reply(message, "I'm here :) :hello-bear:");
    if (message.text == "collect") {
        bot.replyInThread(message, ':waitingmaas: Give me some time to get all the information from Grafana :construction-2:')
        exec('node server.js', (err, stdout, stderr) => {
            console.log(stderr);
            console.log(stdout);
            bot.replyInThread(message, 'All the screenshots is now done, now just type my name and "reports" to receive all the screenshots')
        })
    }

    if (message.text.includes("reports")) {
        // bot.replyInThread(message, 'Here are the reports')
        uploadTheFiles(bot, message, ["ge4Dashboard/ge4-dashboard.png",
            "ge4MonthEnd/ge4Month-Infra.png",
            "ge4MonthEnd/ge4Month-Checklist.png",
            "ge4MonthEnd/ge4Month-Oracle.png",
            "ge4MonthEnd/ge4Month-HanaHa4.png",
            "ge4MonthEnd/ge4MonthEnd-stoBrSoftlayerConnectivity.png",
            "ge4MonthEnd/ge4Month-HortoCoreSwitch.png"
        ])
    }
});

async function uploadTheFiles(bot, message, images) {
    var i = 0;
    while (i < images.length) {
        bot.api.files.upload({
            file: fs.createReadStream(images[i]),
            filename: images[i],
            filetype: "png",
            channels: message.channel,
            thread_ts: message.thread_ts
        }, function (err, res) {
            if (err) {
                console.log("Failed to add file :(", err)
                bot.reply(message, 'Sorry, there has been an error: ' + err)
            }
        })
        i++;
    }
}

slackBot.startRTM();