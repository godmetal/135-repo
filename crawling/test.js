//slack.js 
var Slack = require('slack-node'); 

webhookUri = "https://hooks.slack.com/services/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"; 

slack = new Slack(); slack.setWebhook(webhookUri); 
slack.webhook(
	{
		channel: "#general", // 현 슬랙의 채널 
		username: "test_webhook", // 슬랙에서 보여질 웹훅 이름 
		text: "test" //텍스트 
	}, 
		
	function (err, response) {
		console.log(response); 
	}
);

