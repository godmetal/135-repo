//slack.js 
var Slack = require('slack-node'); 

webhookUri = "https://hooks.slack.com/services/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"; 

slack = new Slack(); slack.setWebhook(webhookUri); 
slack.webhook(
	{
		channel: "#general", // �� ������ ä�� 
		username: "test_webhook", // �������� ������ ���� �̸� 
		text: "test" //�ؽ�Ʈ 
	}, 
		
	function (err, response) {
		console.log(response); 
	}
);
