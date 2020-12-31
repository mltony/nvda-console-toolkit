# nvda-console-toolkit
Console Toolkit is NVDA add-on, that provides accessibility improvements for Windows console, also known as Command prompt. It also works well in Windows PowerShell. Some of the features may work in alternative terminals, such as Cygwin, PuTTY and Windows Terminal, however, the add-on has only been carefully tested with the default Windows Console. SSH users might find this add-on especially handy.

Some of the features were previously part of [Tony's enhancements add-on](https://github.com/mltony/nvda-tonys'enhancements/).

## Downloads

[Console toolkit](https://github.com/mltony/nvda-console-toolkit/releases/latest/download/consoleToolkit.nvda-addon)

## Real-time console speech

This option makes NVDA to speak new lines immediately as they appear in console output, instead of queueing new speech utterances. For example, if NVDA is busy speaking a line that appeared on the screen 1 minute ago, and now a new line appears, this option will cancel speaking the old line and start speaking the new line right away, thus providing a more real-time feedback on what's happening in console window.

## Beep on console updates

Beep a low pitch impulse every time console text is updated.

## Enforce Control+V in consoles

This option makes Control+V shortcut to work inside `ssh` sessions.

## Experimental: command prompt editing

Note: this feature is experimental. Please read  this section carefully and make sure you understand how it works before reporting issues.

Press `NVDA+E` to identify current prompt in console window and edit it in an accessible "Edit prompt" window. After editing you can either press `Escape` to update current command line, or `Enter` to update and immediately execute command. Alternatively you can press `Alt+F4` to close edit prompt window without updating command line.

This feature has been tested in Windows command prompt `cmd.exe` as well as in bash shell over ssh connections, as well as in WSL and cygwin. It might also work in alternative Unix shells, however it hasn't been tested.

Here is how add-on extracts current command.
1. It presses `End` key and then sends a control character, that is a rare Unicodecharacter not likely to be used anywhere.
2. Then it presses `home` key and sends another control character.
3. Then it waits for control characters to appear on the screen, which might take some time on slow SSH connections.
4. Command is what appears between two control characters.
5. When "Use UI Automation to access the Windows Console when available" option is enabled in NVDA settings, it sends one more control character in the beginning of the string. This is needed to parse multiline commands correctly: UIA implementation trims whitespaces in the end of each line, so in order to deduce whether there is a space between two lines, we need to shift them by one character. Please note, however, that this way we don't preserve the number of spaces between words, we only guarantee to preserve the presence of spaces.
6. Before editing add-on makes sure to remove control characters by placing cursor in the beginning and end and simulating `Delete` and `Backspace` key presses.
7. It presents command in "Edit prompt" window for user to view or edit.
8. After user presses `Enter` or `Escape`,it first erases current line in the console.  This is achieved via one of four methods, the choice of the method is configurable. Currently four methods are supported:
    - `Control+C`: works in both `cmd.exe` and `bash`, but leaves previous prompt visible on the screen; doesn't work in emacs; sometimes unreliable on slow SSH connections
    - `Escape`: works only in `cmd.exe`"),
    - `Control+A Control+K`: works in `bash` and `emacs`; doesn't work in `cmd.exe`
    - `Backspace` (recommended): works in all environments; however slower and may cause corruption if the length of the line has changed
9. Then add-on simulates keystrokes to type the updated command and optionally simulates `Enter` key press.

Troubleshooting:
- Verify that 'Home', 'End', 'Delete' and 'Backspace' keys work as expected in your console.
- Verify that your console supports Unicode characters. Some ssh connections don't support Unicode.
- Verify that selected deleting method works in your console.

## Experimental: capturing command output

Note: this feature is experimental. Please read  this section carefully and make sure you understand how it works before reporting issues.

While in command line or in "Edit prompt" window, press `Control+Enter` to capture command output. This add-on is capable of capturing large output that spans multiple screens, although when output is larger than 10 screens capturing process takes significant time to complete. Add-on will play a long chime sound, and it will last as long as the add-on is capturing the output of currently running command, or until timeout has been reached. Alternatively, press `NVDA+E` to interrupt capturing.

When "Use UI Automation to access the Windows Console when available" feature is enabled in NVDA settings, you can switch to other windows while capturing is going on. However, if this option is disabled, then NVDA is using a legacy console code, that only works when consoel is focused, and therefore switching to any other window will pause capturing.

Command capturing works by redirecting command output to `less` command. Default suffix that is appended to commands is:
```
|less -c 2>&1
```
Please only change it if you know what you're doing. This add-on knows how to interact with the output of `less` command to retrieve output page by page.

On Windows `less.exe` tool needs to be installed separately. You can install it via cygwin, or download a windows binary elsewhere.

If you are using `tmux` or `screen` in Linux, please make sure that no status line is displayed in the bottom. In `tmux` run 
```
tmux set status off
```
to get rid of status line, or modify your `tmux.conf` file.

Troubleshooting:
- After a failed output capturing attempt, press `UpArrow` in the console to check what command has actually been executed.
- Revert back to default capturing suffix, mentioned above.
- Try troubleshooting steps from "command prompt editing" section.