<div align="center">

## Microsoft Agent How To \- Basics


</div>

### Description

This code is designed to show you some ofthe basic functions you can perform with Microsoft Agent. Microsoft Agent is that small animal/icon/character that sometimes appears on programs like the Microsoft Office paper clip.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dino Roger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dino-roger.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dino-roger-microsoft-agent-how-to-basics__1-43780/archive/master.zip)





### Source Code

```
'INSTRUCTIONS
'* Add the component "Microsoft Agent Control"
'* Add the comonent "Microsoft Agent Control" to the form
'* Add a text box alled TEXT1
'* Add a button called CMD_HIDE
'* Add a button called CMD_SHOW
'* Add a button called CMD_SPEAK
'* Add a button called CMD_ANIMATION
'* Add a button called CMD_STOP
'* Add a button called CMD_FREEZE
'* Add a button called CMD_SHOWHIDE
'* Add a button called CMD_SPEAKQUESTION
'* Add a list box called LST_ANIMATIONS
'
' You may need to install the merlin agent character if not already installed
'URL to download the character = http://www.msagentring.org/download.asp?char=merlin
'NO DECLARATIONS
'START - CODE
 Private Sub cmd_freeze_Click()
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 merlin.Stop 'Will pause the current animation
 End Sub
 Private Sub cmd_speak_Click()
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 If Not Text1.Text = "" Then merlin.Speak Text1.Text 'Will speak whatever text that is entered in the text box if text is not blank
 End Sub
 Private Sub cmd_Hide_Click()
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 merlin.Hide 'Hides Merlin
 End Sub
 Private Sub cmd_Show_Click()
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 merlin.Show 'Shows Merlin in the screen
 End Sub
 Private Sub cmd_animation_Click()
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 merlin.Play lst_animations.Text 'Will play animation that is selected in the animations list box
 End Sub
 Private Sub cmd_Stop_Click()
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 merlin.Stop 'Will stop the animation - needed for some animations that have built in loop
 merlin.Play "RestPose" 'Will reset Merlin's pose
 End Sub
 Private Sub Cmd_speakquestion_Click()
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 If Not Text1.Text = "" Then merlin.Think Text1.Text 'Will speak whatever text that is entered in the text box if text is not blank. Will display in thought bubble.
 End Sub
 Private Sub Cmd_showhide_Click()
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 merlin.ShowPopupMenu 100, 100 'This will display a clickable HIDE button next to Merlin
 End Sub
 Private Sub Form_Load()
 'You can specify other Microsoft Agents by pointing to a specific agent file.
 'The agent file will have an ACS extension. Remove the quote from the following line and edit.
 'Agent1.Characters.Load "Merlin", "C:\MYAGENT\Merlin.acs"
 Agent1.Characters.Load "merlin", Path 'Needed to access Merlin commands - If you specify your own agent file, delete this line.
 Set merlin = Agent1.Characters("merlin") 'Needed on every object that calls Merlin
 merlin.LanguageID = &H409 'Sets 'Merlin language
 merlin.Show 'Shows Merlin on the screen
 'START - Fill list box with animations
 'Listed this way to show all the animations.
 lst_animations.AddItem "Acknowledge"
 lst_animations.AddItem "Alert"
 lst_animations.AddItem "Announce"
 lst_animations.AddItem "Blink"
 lst_animations.AddItem "Confused"
 lst_animations.AddItem "Congratulate"
 lst_animations.AddItem "Congratulate_2"
 lst_animations.AddItem "Decline"
 lst_animations.AddItem "DoMagic1"
 lst_animations.AddItem "DoMagic2"
 lst_animations.AddItem "DontRecognize"
 lst_animations.AddItem "Explain"
 lst_animations.AddItem "GestureDown"
 lst_animations.AddItem "GestureLeft"
 lst_animations.AddItem "GestureRight"
 lst_animations.AddItem "GestureUp"
 lst_animations.AddItem "GetAttention"
 lst_animations.AddItem "GetAttentionContinued"
 lst_animations.AddItem "GetAttentionReturn"
 lst_animations.AddItem "Greet"
 lst_animations.AddItem "Hearing_1"
 lst_animations.AddItem "Hearing_2"
 lst_animations.AddItem "Hearing_3"
 lst_animations.AddItem "Hearing_4"
 lst_animations.AddItem "Hide"
 lst_animations.AddItem "Idle1_1"
 lst_animations.AddItem "Idle1_2"
 lst_animations.AddItem "Idle1_3"
 lst_animations.AddItem "Idle1_4"
 lst_animations.AddItem "Idle2_1"
 lst_animations.AddItem "Idle2_2"
 lst_animations.AddItem "Idle3_1"
 lst_animations.AddItem "Idle3_2"
 lst_animations.AddItem "LookDown"
 lst_animations.AddItem "LookDownBlink"
 lst_animations.AddItem "LookDownReturn"
 lst_animations.AddItem "LookLeft"
 lst_animations.AddItem "LookLeftBlink"
 lst_animations.AddItem "LookLeft"
 lst_animations.AddItem "LookLeftBlink"
 lst_animations.AddItem "LookLeftReturn"
 lst_animations.AddItem "LookRight"
 lst_animations.AddItem "LookRightBlink"
 lst_animations.AddItem "LookRightReturn"
 lst_animations.AddItem "LookUp"
 lst_animations.AddItem "LookUpBlink"
 lst_animations.AddItem "LookUpReturn"
 lst_animations.AddItem "MoveDown"
 lst_animations.AddItem "MoveLeft"
 lst_animations.AddItem "MoveRight"
 lst_animations.AddItem "MoveUp"
 lst_animations.AddItem "Pleased"
 lst_animations.AddItem "Process"
 lst_animations.AddItem "Processing"
 lst_animations.AddItem "Read"
 lst_animations.AddItem "ReadContinued"
 lst_animations.AddItem "Reading"
 lst_animations.AddItem "ReadReturn"
 lst_animations.AddItem "RestPose"
 lst_animations.AddItem "Sad"
 lst_animations.AddItem "Search"
 lst_animations.AddItem "Searching"
 lst_animations.AddItem "Show"
 lst_animations.AddItem "StartListening"
 lst_animations.AddItem "StopListening"
 lst_animations.AddItem "Suggest"
 lst_animations.AddItem "Surprised"
 lst_animations.AddItem "Think"
 lst_animations.AddItem "Thinking"
 lst_animations.AddItem "Uncertain"
 lst_animations.AddItem "Wave"
 lst_animations.AddItem "Write"
 lst_animations.AddItem "WriteContinued"
 lst_animations.AddItem "WriteReturn"
 lst_animations.AddItem "Writing"
 'END - Fill list box with animations
 'START - Merlin introduction
 merlin.Play "Announce"
 merlin.Play "RestPose"
 merlin.Play "Explain"
 merlin.Speak "Welcome to the Microsoft Agent how to VB6 snippet. Brought to you by http://allvb.net"
 merlin.Play "DoMagic1"
 merlin.Play "DoMagic2"
 merlin.Play "RestPose"
 merlin.Play "Explain"
 merlin.Speak "You can use my code to learn how to use Microsoft Agent in your own programs."
 merlin.Play "Wave"
 merlin.Play "RestPose"
 'END - Merlin introduction
 End Sub
'END - CODE
```

