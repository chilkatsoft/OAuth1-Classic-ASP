<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Finish Quickbooks OAuth1.0a</title>
</head>

<body>

<%
	' The incoming params are  oauth_token and oauth_verifier
	Session("oauth_token") = request.querystring("oauth_token")
	Session("oauth_verifier") = request.querystring("oauth_verifier")
	
	set http = Server.CreateObject("Chilkat_9_5_0.Http")

	success = http.UnlockComponent(Session("chilkatUnlockCode"))
	If (success <> 1) Then
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	End If

	http.OAuth1 = 1
	http.OAuthConsumerKey = Session("consumerKey")
	http.OAuthConsumerSecret = Session("consumerSecret")
	http.OAuthCallback = Session("oauthCallback")
    http.OAuthToken = Session("oauth_token")
    http.OAuthTokenSecret = Session("oauth_token_secret")
    http.OAuthVerifier = Session("oauth_verifier")

	' Exchange the oauth_token for an access token.
	set req = Server.CreateObject("Chilkat_9_5_0.HttpRequest")
    set resp = http.PostUrlEncoded("https://oauth.intuit.com/oauth/v1/get_access_token", req)
	If (resp Is Nothing ) Then
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	ElseIf (resp.StatusCode = 200) Then
		' Success
		set hashTab = Server.CreateObject("Chilkat_9_5_0.Hashtable")
		hashTab.AddQueryParams(resp.BodyStr)
        Session("oauth_token") = hashTab.LookupStr("oauth_token")
        Session("oauth_token_secret") = hashTab.LookupStr("oauth_token_secret")

	    Response.Write "<p>Access Token: " & Session("oauth_token") & "</p>"
	    Response.Write "<p>Access Token Secret: " & Session("oauth_token_secret") & "</p>"
	Else
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	End If

%>


</body>
</html>
