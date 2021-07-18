from O365 import Account

credentials = ('977c94d2-5808-4349-b99a-42a1e629b0ef', 'n79-f-G_0he_4_e303TlvybP2iSMAry4.E')

account = Account(credentials)
m = account.new_message()
m.to.add('p.lerner@fnysllc.com')
m.subject = 'Testing!'
m.body = "George Best quote: I've stopped drinking, but only while I'm asleep."
m.send()