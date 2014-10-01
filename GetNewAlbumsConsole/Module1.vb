Imports System.Net
Imports System.IO
Imports HtmlAgilityPack
Imports System.Text
Imports System.Globalization

Module Module1
    'Where do you want the output file to be saved? (eg...in your 'My Documents' folder)
    Public outputFilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

    'What do you want to name the output file that is created?
    Public outputFileName As String = "AlbumsToDownload.txt"

    Public fullOutputFilePath As String = outputFilePath + "/" + outputFileName

    Sub Main()
        Console.WriteLine("----------STARTING----------")

        Dim inputArgumentMinScore As String = "/minScore="
        Dim inputArgumentMaxScore As String = "/maxScore="
        Dim inputArgumentOutputFilePath As String = "/outputFilePath="

        Try
            'Variables
            Dim theUrl As String = String.Empty
            Dim theHtml As String = String.Empty

            Dim minScore As Integer = 80
            Dim maxScore As Integer = 100

            'Get the user-entered minimum and maximum metascores
            For Each s As String In My.Application.CommandLineArgs
                Console.WriteLine(s)
                If s.StartsWith(inputArgumentMinScore) Then
                    Console.WriteLine("minScore argument passed")
                    minScore = s.Remove(0, inputArgumentMinScore.Length)
                ElseIf s.StartsWith(inputArgumentMaxScore) Then
                    Console.WriteLine("maxScore argument passed")
                    maxScore = s.Remove(0, inputArgumentMaxScore.Length)
                ElseIf s.StartsWith(inputArgumentOutputFilePath) Then
                    Console.WriteLine("outputFilePath argument passed")
                    fullOutputFilePath = s.Remove(0, inputArgumentOutputFilePath.Length)
                End If
            Next

            'Set the URL to get
            theUrl = "http://www.metacritic.com/browse/albums/release-date/new-releases/date?view=detailed"

            Dim web As New HtmlWeb()
            web.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/37.0.2062.124 Safari/537.36"
            Dim doc As HtmlDocument = web.Load(theUrl)

            If IsNothing(doc.DocumentNode.SelectNodes("//p[contains(@class,'no_data')]")) Then
                Dim tags As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div/ol/li/div[contains(@class,'product_wrap')]")

                Dim theArtist As String = ""
                Dim theAlbum As String = ""
                Dim theLink As String = ""
                Dim theMetascore As String = ""
                Dim theReleaseDate As String = ""
                Dim theGenres As String = ""
                Dim theImage As String = ""
                'Dim theArtwork As Image
                Dim theAlbumData As String = ""
                Dim theAlbumCompareString As String = ""

                'Debug.WriteLine("ITEMS FOUND: [" + tags.Count.ToString + "]")

                If Not IsNothing(tags) Then
                    For Each aTag In tags
                        Dim nodeAlbum As HtmlNode = aTag.SelectSingleNode(".//h3[@class='product_title']/a")
                        Dim nodeArtist As HtmlNode = aTag.SelectSingleNode(".//h3[@class='product_title']/span[@class='product_artist']")
                        Dim nodeMetascore As HtmlNode = aTag.SelectSingleNode(".//a[contains(@class,'product_score')]/span")
                        Dim nodeReleaseDate As HtmlNode = aTag.SelectSingleNode(".//li[contains(@class,'release_date')]/span[contains(@class,'data')]")
                        Dim nodeGenres As HtmlNode = aTag.SelectSingleNode(".//li[contains(@class,'genre')]/span[contains(@class,'data')]")
                        Dim nodeImage As HtmlNode = aTag.SelectSingleNode(".//img[contains(@class,'product_image')]")

                        'Debug.WriteLine("    ITEMS FOUND: [" + aTag.SelectSingleNode(".//h3[@class='product_title']/a").InnerHtml + "]")

                        If Not IsNothing(nodeAlbum) Then
                            'Debug.WriteLine("ALBUM: " + nodeAlbum.InnerHtml)
                            'Debug.WriteLine("LINK: http://www.metacritic.com/" + nodeAlbum.Attributes("href").Value.ToString)

                            theAlbum = nodeAlbum.InnerHtml
                            theLink = "http://www.metacritic.com/" + nodeAlbum.Attributes("href").Value.ToString
                        End If

                        If Not IsNothing(nodeArtist) Then
                            'Debug.WriteLine("ARTIST: " + nodeArtist.InnerHtml.Replace(" - ", ""))

                            theArtist = nodeArtist.InnerHtml.Replace(" - ", "")
                        End If

                        If Not IsNothing(nodeMetascore) Then
                            'Debug.WriteLine("METASCORE: " + nodeMetascore.InnerHtml)

                            theMetascore = nodeMetascore.InnerHtml
                        End If

                        If Not IsNothing(nodeReleaseDate) Then
                            'Debug.WriteLine("RELEASE DATE: " + nodeReleaseDate.InnerHtml)

                            theReleaseDate = nodeReleaseDate.InnerHtml
                        End If

                        If Not IsNothing(nodeGenres) Then
                            'Debug.WriteLine("GENRES: " + nodeGenres.InnerHtml.Replace("  ", "").TrimStart.TrimEnd)

                            theGenres = nodeGenres.InnerHtml.Replace("  ", "").TrimStart.TrimEnd
                        End If

                        If Not IsNothing(nodeImage) Then
                            'Debug.WriteLine("IMAGE: " + nodeImage.Attributes("src").Value.ToString)

                            theImage = nodeImage.Attributes("src").Value.ToString

                            'Dim myWebClient As New System.Net.WebClient
                            'Dim imageInBytes() As Byte = myWebClient.DownloadData(theImage)
                            'Dim imageStream As New IO.MemoryStream(imageInBytes)
                            'theArtwork = New System.Drawing.Bitmap(imageStream)
                            'Dim imageListSmall As New ImageList()
                            'imageListSmall.Images.Add(artwork)
                            'PictureBox1.Image = New System.Drawing.Bitmap(imageStream)
                            'lvAlbums.SmallImageList = imageListSmall
                        End If

                        If CType(theMetascore, Integer) >= minScore AndAlso CType(theMetascore, Integer) <= maxScore Then
                            'Debug.WriteLine(theArtist + "|" + theAlbum + "|" + theMetascore + "|" + theReleaseDate)
                            'tbResults.AppendText(theArtist + "|" + theAlbum + "|" + theMetascore + "|" + theReleaseDate + vbCrLf)

                            'Dim aRow As New DataGridViewRow
                            'aRow.CreateCells(dgvAlbums)
                            'aRow.Height = 60
                            'aRow.Cells(0).Value = theArtwork
                            'aRow.Cells(1).Value = theMetascore
                            'aRow.Cells(2).Value = theAlbum
                            'aRow.Cells(3).Value = theArtist
                            'aRow.Cells(4).Value = theReleaseDate
                            'aRow.Cells(5).Value = theGenres

                            'aRow.Cells(1).Style.ForeColor = Color.White
                            'dgvAlbums.Columns(1).DefaultCellStyle.Font = New Font("Helvetica Neue", 16, FontStyle.Bold)
                            'dgvAlbums.Columns(1).DefaultCellStyle.Alignment = ContentAlignment.MiddleCenter

                            'If theMetascore >= 70 Then
                            '    aRow.Cells(1).Style.BackColor = Color.LimeGreen
                            'ElseIf theMetascore >= 50 Then
                            '    aRow.Cells(1).Style.BackColor = Color.Yellow
                            'Else
                            '    aRow.Cells(1).Style.BackColor = Color.Red
                            'End If

                            ''dgvAlbums.Rows(dgvAlbums.Rows.Count - 1).Cells("colMetascore").Value = theMetascore

                            'dgvAlbums.Rows.Add(aRow)


                            theAlbumCompareString = "<title>" + theArtist + " " + theAlbum + "</title>"


                            theAlbumData = "<artist>" + theArtist + "</artist>" +
                                "<album>" + theAlbum + "</album>" +
                                "<metascore>" + theMetascore + "</metascore>" +
                                "<releaseDate>" + theReleaseDate + "</releaseDate>" +
                                "<genres>" + theGenres + "</genres>" +
                                "<link>" + theLink + "</link>" +
                                "<imageLink>" + theImage + "</imageLink>" +
                                "<title>" + theArtist + " " + theAlbum + "</title>"

                            AddAlbumToTextFile(theAlbumCompareString, theAlbumData, theArtist, theAlbum, theMetascore.ToString)
                        End If
                    Next
                End If

                'Remove the trailing vbCrLf from the results
                'tbResults.Text = tbResults.Text.TrimEnd(vbCrLf)

            Else
                Throw New Exception("Page returned 'No Results Found' message")
            End If
        Catch ex As Exception
            'lblErrorMessage.Visible = True
            'lblErrorMessage.Text = lblErrorMessage.Text + vbCrLf + ex.Message.ToString
            Console.WriteLine(ex.Message.ToString)
        Finally
            'btnProcess.Text = "Process"
            'btnProcess.Enabled = True
            Console.WriteLine("----------FINISHED----------")
        End Try
    End Sub

    Public Sub AddAlbumToTextFile(ByVal albumComparisonString As String, ByVal albumData As String, ByVal theArtist As String, ByVal theAlbum As String, ByVal theMetascore As String)
        Dim path As String = fullOutputFilePath

        ' This text is added only once to the file. 
        If File.Exists(path) = False Then
            ' Create a file to write to. 
            Dim createText As String = ""
            File.WriteAllText(path, createText)
        End If

        ' Open the file to read from. 
        Dim readText As String = File.ReadAllText(path)


        'Check if album is already in file
        If Not readText.Contains(RemoveDiacritics(albumData)) Then
            Console.WriteLine("   Adding: [" + theArtist + " - " + theAlbum + " (" + theMetascore + ")]")

            ' This text is always added, making the file longer over time 
            ' if it is not deleted. 
            Dim appendText As String = RemoveDiacritics(albumData) + Environment.NewLine
            File.AppendAllText(path, appendText)
        End If
    End Sub

    Private Function RemoveDiacritics(ByVal text As String) As String
        Dim normalizedString = text.Normalize(NormalizationForm.FormD)
        Dim stringBuilder = New StringBuilder()

        For Each c In normalizedString
            Dim unicodeCategory__1 = CharUnicodeInfo.GetUnicodeCategory(c)
            If unicodeCategory__1 <> UnicodeCategory.NonSpacingMark Then
                stringBuilder.Append(c)
            End If
        Next

        Return stringBuilder.ToString().Normalize(NormalizationForm.FormC)
    End Function
End Module
