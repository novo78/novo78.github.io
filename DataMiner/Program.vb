Imports System

Module Program
    Sub Main(args As String())

        'Read in the Garmin File:

        Dim objTextReader As New IO.StreamReader("C:\Users\RichardBristow\Documents\python-garminconnect\output.txt")
        Dim strGarmin As String = objTextReader.ReadToEnd
        objTextReader.Close()

        Dim strGarData As String = String.Empty

        For Each garline As String In strGarmin.Split(vbLf)
            If garline.Contains("{") Then
                'Console.WriteLine(garline & vbCrLf)
                strGarData += garline.Replace("{", "").Replace("}", "").Replace("]", "").Replace("[", "")
            End If
        Next

        Dim intRestingCalories, intCaloriesBurnt, intActiveCalories, intCaloriesConsumed, intCalorieGoal, intRemainingCalories As Integer
        Dim intAdjustedGoal, intDailyRemainingCalories As Integer
        Dim intTotalSteps, intStepGoal, intTotalDistance As Integer
        Dim intSleepStart, intSleepEnd, intDeepSleep, intLightSleep, intREMSleep, intAwake As Int64

        Dim DatalineOut As String = "C:\Users\RichardBristow\Documents\python-garminconnect\dataline.txt"
        Dim DatalineRawOut As String = "C:\Users\RichardBristow\Documents\python-garminconnect\dataline_raw.txt"
        Dim Dawnoftime As New DateTime
        Dawnoftime = DateTime.Parse("January 1 1970 12:00:00 am")

        'Output to Text Files
        Dim bolWriteOut As Boolean = False

        'File Clean up
        If FileIO.FileSystem.FileExists(DatalineOut) Then FileIO.FileSystem.DeleteFile(DatalineOut)
        If FileIO.FileSystem.FileExists(DatalineRawOut) Then FileIO.FileSystem.DeleteFile(DatalineRawOut)

        For Each dataline As String In strGarData.Split(", ")

            If bolWriteOut = True Then FileIO.FileSystem.WriteAllText(DatalineRawOut, dataline & vbCrLf, True)

            If dataline.ToLower.Contains("none") Then
                dataline = dataline.ToLower.Replace("none", CInt(0))
            End If

            If bolWriteOut = True Then FileIO.FileSystem.WriteAllText(DatalineOut, dataline & vbCrLf, True)

            If dataline.Contains("'bmrKilocalories':") Then
                intRestingCalories = dataline.Replace("'bmrKilocalories': ", "")
            ElseIf dataline.Contains("'totalKilocalories':") Then
                intCaloriesBurnt = dataline.Replace("'totalKilocalories': ", "")
            ElseIf dataline.Contains("'activeKilocalories':") Then
                intActiveCalories = dataline.Replace("'activeKilocalories': ", "")
            ElseIf dataline.Contains("'consumedKilocalories':") Then
                If dataline.ToLower.Contains("none") Then
                    intCaloriesConsumed = 0
                Else
                    intCaloriesConsumed = dataline.Replace("'consumedKilocalories': ", "")
                End If
            ElseIf dataline.Contains("'netCalorieGoal':") Then
                intCalorieGoal = dataline.Replace("'netCalorieGoal': ", "")
            ElseIf dataline.Contains("'netRemainingKilocalories':") Then
                intRemainingCalories = dataline.Replace("'netRemainingKilocalories': ", "")

                'Steps Data
            ElseIf dataline.Contains("'totalSteps':") Then
                intTotalSteps = dataline.Replace("'totalSteps': ", "")
            ElseIf dataline.Contains("'dailyStepGoal':") Then
                intStepGoal = dataline.Replace("'dailyStepGoal': ", "")
            ElseIf dataline.Contains("'totalDistanceMeters':") Then
                intTotalDistance = dataline.Replace("'totalDistanceMeters': ", "")

                'Sleep Data
            ElseIf dataline.Contains("'autoSleepStartTimestampGMT':") Then
                dataline = dataline.Substring(0, dataline.Length - 3)
                intSleepStart = dataline.Replace("'autoSleepStartTimestampGMT': ", "")
            ElseIf dataline.Contains("'autoSleepEndTimestampGMT':") Then
                dataline = dataline.Substring(0, dataline.Length - 3)
                intSleepEnd = dataline.Replace("'autoSleepEndTimestampGMT': ", "")
            ElseIf dataline.Contains("'deepSleepSeconds':") Then
                intDeepSleep = dataline.Replace("'deepSleepSeconds': ", "")
            ElseIf dataline.Contains("'lightSleepSeconds':") Then
                intLightSleep = dataline.Replace("'lightSleepSeconds': ", "")
            ElseIf dataline.Contains("'remSleepSeconds':") Then
                intREMSleep = dataline.Replace("'remSleepSeconds': ", "")
            ElseIf dataline.Contains("'awakeSleepSeconds':") Then
                intAwake = dataline.Replace("'awakeSleepSeconds': ", "")
            End If

        Next

        'Adjusted Goal
        intAdjustedGoal = intCalorieGoal + intActiveCalories

        'Daily Remaining
        intDailyRemainingCalories = intAdjustedGoal - intCaloriesConsumed

        Console.WriteLine("Daily Summmary")
        Console.WriteLine(vbCr)
        Console.WriteLine("Calorie Data")
        Console.WriteLine("Calorie Goal = " & intCalorieGoal)
        Console.WriteLine("Resting Calories = " & intRestingCalories)
        Console.WriteLine("Calories Burnt = " & intCaloriesBurnt)
        Console.WriteLine("Active Calories = " & intActiveCalories)
        Console.WriteLine("Calories Consumed = " & intCaloriesConsumed)
        Console.WriteLine("Remaining Calories = " & intRemainingCalories)
        Console.WriteLine("Adjusted Goal = " & intAdjustedGoal)
        Console.WriteLine("Daily Remaining Calories = " & intDailyRemainingCalories)
        Console.WriteLine(vbCr)
        Console.WriteLine("Steps Data")
        Console.WriteLine("Step Goal = " & intStepGoal)
        Console.WriteLine("Total Steps = " & intTotalSteps)
        Console.WriteLine("Total Meters = " & intTotalDistance)
        Console.WriteLine(vbCr)
        Console.WriteLine("Sleep Data")
        Console.WriteLine("Sleep Start = " & Dawnoftime.AddSeconds(intSleepStart))
        Console.WriteLine("Sleep End = " & Dawnoftime.AddSeconds(intSleepEnd))

        Dim tmDeepSleep As New TimeSpan(TimeSpan.TicksPerSecond * intDeepSleep)
        Dim strDeepSleep As String = tmDeepSleep.Hours.ToString("00") & ":" & tmDeepSleep.Minutes.ToString("00")
        Console.WriteLine("Deep Sleep = " & strDeepSleep)

        Dim tmLightSleep As New TimeSpan(TimeSpan.TicksPerSecond * intLightSleep)
        Dim strLightSleep As String = tmLightSleep.Hours.ToString("00") & ":" & tmLightSleep.Minutes.ToString("00")
        Console.WriteLine("Light Sleep = " & strLightSleep)

        Dim tmREMSleep As New TimeSpan(TimeSpan.TicksPerSecond * intREMSleep)
        Dim strREMSleep As String = tmREMSleep.Hours.ToString("00") & ":" & tmREMSleep.Minutes.ToString("00")
        Console.WriteLine("REM Sleep = " & strREMSleep)

        Dim tmTotalSleep As New TimeSpan
        tmTotalSleep = tmTotalSleep.Add(tmDeepSleep).Add(tmREMSleep).Add(tmLightSleep)
        Dim strTotalSleep As String = tmTotalSleep.Hours.ToString("00") & ":" & tmTotalSleep.Minutes.ToString("00")
        Console.WriteLine("Total Sleep = " & strTotalSleep)

        Dim tmAwake As New TimeSpan(TimeSpan.TicksPerSecond * intAwake)
        Dim strAwake As String = tmAwake.Hours.ToString("00") & ":" & tmAwake.Minutes.ToString("00")
        Console.WriteLine("Awake = " & strAwake)





    End Sub
End Module
