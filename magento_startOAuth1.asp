<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Start Magento OAuth1.0a</title>
</head>

<body>

 
<%
	Session("consumerKey") = "MAGENTO_CONSUMER_KEY"
	Session("consumerSecret") = "MAGENTO_CONSUMER_SECRET"
	Session("oauthCallback") = "http://localhost/magento_finishOAuth1.asp"
	Session("chilkatUnlockCode") = "Anything for 30-day trial"
	
	Session("requestTokenUrl") = "http://www.example.com/oauth/initiate"
    Session("authorizeTokenUrl") = "http://www.example.com/oauth/authorize"
    Session("accessTokenUrl") = "http://www.example.com/oauth/token"
	
	set http = Server.CreateObject("Chilkat_9_5_0.Http")

	success = http.UnlockComponent(Session("chilkatUnlockCode"))
	If (success <> 1) Then
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	End If


	http.OAuth1 = 1
	http.OAuthConsumerKey = Session("consumerKey")
	http.OAuthConsumerSecret = Session("consumerSecret")
	http.OAuthCallback = Session("oauthCallback")

	set req = Server.CreateObject("Chilkat_9_5_0.HttpRequest")
    set resp = http.PostUrlEncoded(Session("requestTokenUrl"), req)
	If (resp Is Nothing ) Then
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	ElseIf (resp.StatusCode = 200) Then
		' Success
		set hashTab = Server.CreateObject("Chilkat_9_5_0.Hashtable")
		hashTab.AddQueryParams(resp.BodyStr)
		'Response.Write "BodyStr: [" & resp.BodyStr & "]<p>"
        Session("oauth_token") = hashTab.LookupStr("oauth_token")
        Session("oauth_token_secret") = hashTab.LookupStr("oauth_token_secret")
		'Response.Write "oauth_token: [" & Session("oauth_token") & "]<p>"
		'Response.Write "oauth_token_secret: [" & Session("oauth_token_secret") & "]<p>"

	    Response.Redirect Session("authorizeTokenUrl") + "?oauth_token=" + Session("oauth_token")
	Else
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	End If

%>


</body>
</html>
