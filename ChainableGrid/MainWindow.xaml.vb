Imports System.Xml
Imports System.Reflection
Imports System.ComponentModel
Imports System.Threading
Imports System.Windows.Markup
Imports <xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
Imports <xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
Imports System.Threading.Tasks

Class MainWindow
    Private WithEvents doc As FixedDocumentReport

    Private WithEvents tmrFirstPage As New System.Timers.Timer

    ' Button Enabling
    Private Sub DisableReportButtons()
        vwr.IsEnabled = False

        barProgress.Height = 30
        barProgress.Visibility = Windows.Visibility.Visible

        btn1.IsEnabled = False
        btn2.IsEnabled = False
        btn3.IsEnabled = False
        btn4.IsEnabled = True
    End Sub
    Private Sub EnableReportButtons()
        btn1.IsEnabled = True
        btn2.IsEnabled = True
        btn3.IsEnabled = True
        btn4.IsEnabled = False

        barProgress.Height = 0
        barProgress.Visibility = Windows.Visibility.Hidden

        tmrFirstPage.Enabled = True
        tmrFirstPage.Interval = 100
    End Sub

    Class Project
        Property Name As String
        Property ProjectNumber As Integer
        Property Stage As String
        Property AnnualOperatingAmount As Decimal
        Property PriorityScore As Integer
    End Class

    Dim prjs As List(Of Project) = New List(Of Project) From
        {
            New Project With {.Name = "Test 3", .ProjectNumber = 321, .Stage = "Inactive", .AnnualOperatingAmount = 1000.0, .PriorityScore = 3},
            New Project With {.Name = "Test 1", .ProjectNumber = 123, .Stage = "Active", .AnnualOperatingAmount = 1000.0, .PriorityScore = 1},
            New Project With {.Name = "Test 2", .ProjectNumber = 231, .Stage = "Construction", .AnnualOperatingAmount = 1000.0, .PriorityScore = 2}
        }

    ' Custom Report
    Private Sub Button_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        DisableReportButtons()

        Dim reportHeader = <Border VerticalAlignment="Top" HorizontalAlignment="Center" Height="80">
                               <StackPanel>
                                   <TextBlock TextAlignment="Center" FontSize="18" FontWeight="Bold" Foreground="DarkBlue">WPF Fixed Document Reports</TextBlock>
                                   <TextBlock TextAlignment="Center" FontSize="18" FontWeight="Bold" Foreground="DarkBlue">Example Project List Report for FY<%= Date.Now.Year %></TextBlock>
                               </StackPanel>
                           </Border>
        Dim pageHeader = <Border VerticalAlignment="Top" HorizontalAlignment="Stretch">
                             <Grid TextBlock.FontSize="12" Background="LightGray">
                                 <Grid.ColumnDefinitions>
                                     <ColumnDefinition Width="320"></ColumnDefinition>
                                     <ColumnDefinition Width="180"></ColumnDefinition>
                                     <ColumnDefinition Width="90"></ColumnDefinition>
                                     <ColumnDefinition Width="130"></ColumnDefinition>
                                 </Grid.ColumnDefinitions>
                                 <Border Grid.Column="0" BorderBrush="Black" BorderThickness="1.5">
                                     <TextBlock TextAlignment="Center" VerticalAlignment="Center" Grid.Column="0" TextWrapping="Wrap" FontWeight="Bold">Project Name</TextBlock>
                                 </Border>
                                 <Border Grid.Column="1" BorderBrush="Black" BorderThickness="1.5">
                                     <TextBlock TextAlignment="Center" VerticalAlignment="Center" Grid.Column="1" FontWeight="Bold">Stage</TextBlock>
                                 </Border>
                                 <Border Grid.Column="2" BorderBrush="Black" BorderThickness="1.5">
                                     <TextBlock TextAlignment="Center" VerticalAlignment="Center" Grid.Column="2" FontWeight="Bold">Project Number</TextBlock>
                                 </Border>
                                 <Border Grid.Column="3" BorderBrush="Black" BorderThickness="1.5">
                                     <TextBlock TextAlignment="Center" VerticalAlignment="Center" Grid.Column="3" TextWrapping="Wrap" FontWeight="Bold">Annual Operating Amount</TextBlock>
                                 </Border>
                             </Grid>
                         </Border>
        Dim items = From prj In prjs Select
                    <Border VerticalAlignment="Top" HorizontalAlignment="Stretch">
                        <Grid TextBlock.FontSize="12" Opacity=<%= If(prj.Stage = "Inactive", 1.0, 1.0) %>>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="320"></ColumnDefinition>
                                <ColumnDefinition Width="180"></ColumnDefinition>
                                <ColumnDefinition Width="90"></ColumnDefinition>
                                <ColumnDefinition Width="130"></ColumnDefinition>
                            </Grid.ColumnDefinitions>
                            <Border Grid.Column="0" BorderBrush="Black" BorderThickness="0.5">
                                <TextBlock TextAlignment="Left" TextWrapping="Wrap"><%= prj.Name %></TextBlock>
                            </Border>
                            <!-- Example of conditonal attribute -->
                            <Border Grid.Column="1" BorderBrush="Black" BorderThickness="0.5" Background=<%= If(prj.Stage = "Construction", Brushes.LightYellow, Brushes.White) %>>
                                <TextBlock TextAlignment="Center" FontWeight=<%= If(prj.Stage = "Complete", FontWeights.Bold, FontWeights.Normal) %>><%= prj.Stage %></TextBlock>
                            </Border>
                            <Border Grid.Column="2" BorderBrush="Black" BorderThickness="0.5">
                                <TextBlock TextAlignment="Center"><%= prj.ProjectNumber %></TextBlock>
                            </Border>
                            <Border Grid.Column="3" BorderBrush="Black" BorderThickness="0.5">
                                <TextBlock Margin="5 0" TextAlignment="Left"><%= FormatCurrency(prj.AnnualOperatingAmount, 0) %></TextBlock>
                            </Border>
                        </Grid>
                    </Border>

        doc = New FixedDocumentReport With {.ReportHeader = reportHeader, .PageHeader = pageHeader, .Items = items,
                                                .PageWidth = 816, .PageHeight = 1056, .PageMargin = New Thickness(48)}

        doc.GenerateReport()
        'vwr.Document = doc.GenerateReport()
    End Sub

    ' Standard Table Report
    Private Sub Button2_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        DisableReportButtons()

        FixedDocumentReport.TableReport(
            "Example Project List Report for FY" & Date.Now.Year,
            (From a In prjs Select a.Name, Stage = <TextBlock TextWrapping="Wrap" TextAlignment="Center" Background=<%= If(a.Stage = "Construction", Brushes.LightYellow, Brushes.White) %>><%= a.Stage %></TextBlock>, a.ProjectNumber, AnnualOperatingAmount = FormatCurrency(a.AnnualOperatingAmount, 0), a.PriorityScore).AsEnumerable,
            {4, 3, 1, 1, 1}, {"Project Name", "Stage", "Project Number", "Annual Operating Amount", "Priority Score"},
            {TextAlignment.Left, TextAlignment.Center, TextAlignment.Left, TextAlignment.Right, TextAlignment.Left})

        doc = FixedDocumentReport._currentDoc
    End Sub

    ' Standard Table Report - Long
    Private Sub Button3_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        '    DisableReportButtons()

        '    FixedDocumentReport.TableReport(
        '        "Invoices Report",
        '        (From a In db.Invoices Select btn = <Image Height="50" Source=<%= If(a.ActualAmount > 0, "/Images/Book_Green.ico", "/Images/Book_JournalwPen.ico") %>></Image>, a.ProjectNo, a.ItemHeading, AnnualOperatingAmount = FormatCurrency(a.ActualAmount, 0)).AsEnumerable,
        '        {1, 2, 3, 1}, {"Button", "Project Number", "Invoice Information", "Actual Amount"},
        '        {TextAlignment.Left, TextAlignment.Left, TextAlignment.Left, TextAlignment.Right})

        '    doc = FixedDocumentReport._currentDoc
    End Sub


    Private Sub doc_ReportCompleted() Handles doc.ReportCompleted
        barProgress.Dispatcher.Invoke(
            New Action(
                Sub()
                    EnableReportButtons()
                End Sub))
    End Sub
    Private Sub doc_ReportProgress(itemCount As Integer, currentItem As Integer) Handles doc.ReportProgress
        barProgress.Dispatcher.Invoke(
            New Action(
                Sub()
                    barProgress.Minimum = 0
                    barProgress.Maximum = itemCount
                    barProgress.Value = currentItem
                End Sub))
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        FixedDocumentReport.CancelReport()
        EnableReportButtons()
    End Sub
    Private Sub Button_Click_1(sender As System.Object, e As System.Windows.RoutedEventArgs)
        vwr.FirstPage()
    End Sub
    Private Sub Button_Click_4(sender As System.Object, e As System.Windows.RoutedEventArgs)
        vwr.LastPage()
    End Sub
    Private Sub Button_Click_2(sender As System.Object, e As System.Windows.RoutedEventArgs)
        vwr.PreviousPage()
    End Sub
    Private Sub Button_Click_3(sender As System.Object, e As System.Windows.RoutedEventArgs)
        vwr.NextPage()
    End Sub

    Private Sub tmrFirstPage_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs) Handles tmrFirstPage.Elapsed
        vwr.Dispatcher.Invoke(
             New Action(
                Sub()
                    vwr.Document = doc._doc
                    vwr.FirstPage()
                    vwr.IsEnabled = True
                End Sub))
        tmrFirstPage.Enabled = False
    End Sub
End Class

Public Class FixedDocumentReport
    Private Class StackPanelFlow
        Inherits StackPanel

        Public ChildPanel As StackPanelFlow

        Public Sub NewAddItem(item As FrameworkElement)
            Dim currentHeight = (From a As FrameworkElement In Me.Children Select a.ActualHeight + a.Margin.Top + a.Margin.Bottom).Sum
            item.Arrange(New Rect(0, 0, Me.ActualWidth, Me.ActualHeight - currentHeight))

            If currentHeight + item.ActualHeight + item.Margin.Top + item.Margin.Bottom < Me.ActualHeight Then
                Me.Children.Add(item)
            Else
                If ChildPanel IsNot Nothing Then
                    ChildPanel.Children.Add(item)
                End If
            End If
        End Sub
    End Class

    Private _pnlCurrent As StackPanelFlow
    Private _pnlNext As StackPanelFlow
    Private _pg1 As New PageContent
    Private _pg1Content As FixedPage
    Private _reportHeader As XElement
    Private _pageHeader As XElement
    Private _items As IEnumerable(Of XElement)
    Private _pageWidth As Single
    Private _pageHeight As Single
    Private _pageMargin As Thickness
    Public _doc As FixedDocument

    Private footerHeight = 20

    Public Property ReportHeader() As XElement
        Get
            Return _reportHeader
        End Get
        Set(ByVal value As XElement)
            _reportHeader = value
        End Set
    End Property
    Public Property PageHeader() As XElement
        Get
            Return _pageHeader
        End Get
        Set(ByVal value As XElement)
            _pageHeader = value
        End Set
    End Property
    Public Property Items() As IEnumerable(Of XElement)
        Get
            Return _items
        End Get
        Set(ByVal value As IEnumerable(Of XElement))
            _items = value
        End Set
    End Property
    Public Property PageWidth() As Single
        Get
            Return _pageWidth
        End Get
        Set(ByVal value As Single)
            _pageWidth = value
        End Set
    End Property
    Public Property PageHeight() As Single
        Get
            Return _pageHeight
        End Get
        Set(ByVal value As Single)
            _pageHeight = value
        End Set
    End Property
    Public Property PageMargin() As Thickness
        Get
            Return _pageMargin
        End Get
        Set(ByVal value As Thickness)
            _pageMargin = value
        End Set
    End Property

    Public Function GenerateReport() As FixedDocument
        _currentDoc = Me
        _doc = New FixedDocument

        Dim footer =
                <Border DockPanel.Dock="Bottom" VerticalAlignment="Top" HorizontalAlignment="Left" Height=<%= footerHeight %> Width=<%= _pageWidth - _pageMargin.Left - _pageMargin.Right %>
                    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="7*"></ColumnDefinition>
                            <ColumnDefinition Width="1*"></ColumnDefinition>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Foreground="DarkBlue" FontStyle="Italic"><%= Now.ToLongDateString %></TextBlock>
                        <TextBlock Grid.Column="1" Foreground="DarkBlue" FontStyle="Italic" TextAlignment="Right">Page 1</TextBlock>
                    </Grid>
                </Border>

        _pg1Content = New FixedPage With {.Width = _pageWidth, .Height = _pageHeight}
        _pnlCurrent = New StackPanelFlow With {.Width = _pageWidth - _pageMargin.Left - _pageMargin.Right, .Height = _pageHeight - _pageMargin.Top - _pageMargin.Bottom - footerHeight}
        _pnlNext = New StackPanelFlow With {.Width = _pageWidth - _pageMargin.Left - _pageMargin.Right, .Height = _pageHeight - _pageMargin.Top - _pageMargin.Bottom - footerHeight}

        Dim dp As New DockPanel With {.Margin = New Thickness(48)}
        dp.Children.Add(XamlReader.Load(footer.CreateReader))
        dp.Children.Add(_pnlCurrent)

        ' Set the relationship between the current and next panel so that items will flow to the next panel
        _pnlCurrent.ChildPanel = _pnlNext

        ' Add the current page to the first page and add the first page to the collection of pages
        _pg1Content.Children.Add(dp)
        _pg1.Child = _pg1Content
        _doc.Pages.Add(_pg1)

        ' Arrange the panels and set the page header
        With _pnlCurrent
            .Arrange(New Rect(0, 0, .Width, .Height))
            _pnlNext.Arrange(New Rect(0, 0, _pnlCurrent.Width, _pnlCurrent.Height))

            ' Add the report header
            If _reportHeader IsNot Nothing Then
                Dim reportHeader = XamlReader.Load(_reportHeader.CreateReader)
                reportHeader.Arrange(New Rect(0, 0, .Width, .Height))
                .Children.Add(reportHeader)
            End If

            ' Add the page header
            If _pageHeader IsNot Nothing Then
                Dim hdr = XamlReader.Load(_pageHeader.CreateReader)
                hdr.Arrange(New Rect(0, 0, .ActualWidth, .ActualHeight))
                .NewAddItem(hdr)
            End If
        End With

        ' Add all of the items to the report causing a cascade of additional pages, if necessary
        AddItemsAsync(_items)

        Return _doc
    End Function
    Public Sub AddItem(item As XElement)
        ' Add the item the current page
        _pnlCurrent.NewAddItem(XamlReader.Load(item.CreateReader))

        ' If the item instead was put on the next page, then make a new page with the next panel
        If _pnlNext.Children.Count > 0 Then
            Dim footer =
                <Border DockPanel.Dock="Bottom" VerticalAlignment="Top" HorizontalAlignment="Left" Height=<%= footerHeight %> Width=<%= _pageWidth - _pageMargin.Left - _pageMargin.Right %>
                    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="7*"></ColumnDefinition>
                            <ColumnDefinition Width="1*"></ColumnDefinition>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Foreground="DarkBlue" FontStyle="Italic"><%= Now.ToLongDateString %></TextBlock>
                        <TextBlock Grid.Column="1" Foreground="DarkBlue" FontStyle="Italic" TextAlignment="Right">Page <%= _currentDoc._doc.Pages.Count + 1 %></TextBlock>
                    </Grid>
                </Border>

            Dim dp As New DockPanel With {.Margin = New Thickness(48)}
            dp.Children.Add(XamlReader.Load(footer.CreateReader))
            dp.Children.Add(_pnlNext)

            _pg1 = New PageContent
            _pg1Content = New FixedPage With {.Width = _pageWidth, .Height = _pageHeight}
            _pg1Content.Children.Add(dp)
            _pg1.Child = _pg1Content
            _doc.Pages.Add(_pg1)

            _pnlCurrent = _pnlNext
            With _pnlCurrent
                _pnlNext = New StackPanelFlow With {.Width = _pnlCurrent.Width, .Height = _pnlCurrent.Height}
                _pnlNext.Arrange(New Rect(0, 0, .Width, .Height))
                .ChildPanel = _pnlNext

                ' If the page is not brought into view, then it will have zero actual height and width
                _pg1Content.BringIntoView()

                ' Remove the item from the page, add the page header, and add the item back
                Dim temp = .Children(0)
                .Children.Clear()
                If _pageHeader IsNot Nothing Then
                    Dim hdr = XamlReader.Load(_pageHeader.CreateReader)
                    hdr.Arrange(New Rect(0, 0, _pnlCurrent.ActualWidth, _pnlCurrent.ActualHeight))
                    .NewAddItem(hdr)
                End If
                .NewAddItem(temp)
            End With
        End If
    End Sub
    Public Sub AddItems(items As IEnumerable(Of XElement))
        For Each item In items
            AddItem(item)
        Next
    End Sub

    Shared WithEvents wk As BackgroundWorker
    Shared itemCount As Integer
    Shared currentItem As Integer
    Shared Sub AddItemsAsync(items As IEnumerable(Of XElement))
        itemCount = items.Count
        currentItem = 0
        wk = New BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        wk.RunWorkerAsync(items)
    End Sub
    Shared Sub wk_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles wk.DoWork
        Dim items As IEnumerable(Of XElement) = e.Argument

        For Each item In items
            If e.Cancel Then Exit For
            wk.ReportProgress(0, item)
            Thread.Sleep(1)
            RaiseEvent ReportProgress(itemCount, currentItem)
            If wk.CancellationPending Then Exit For
            currentItem += 1
        Next
    End Sub
    Shared Sub wk_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles wk.ProgressChanged
        Dim item As XElement = e.UserState

        If _currentDoc IsNot Nothing Then
            _currentDoc.AddItem(item)
        End If
    End Sub
    Private Shared Sub wk_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles wk.RunWorkerCompleted
        RaiseEvent ReportCompleted()
    End Sub
    Public Shared Sub CancelReport()
        If wk IsNot Nothing AndAlso Not wk.CancellationPending Then
            wk.CancelAsync()
        End If
        _currentDoc = Nothing
    End Sub

    Public Shared _currentDoc As FixedDocumentReport
    Public Shared Function TableReport(headerText As String, tbl As IEnumerable, columnWidths() As Double, columnHeader() As String, columnAlignments() As TextAlignment) As FixedDocument
        Dim info() As PropertyInfo = (From a In tbl).First.GetType().GetProperties()
        Dim propertyMap As New SortedList(Of String, Integer)
        For Each prop In info
            propertyMap.Add(prop.Name, propertyMap.Count)
        Next

        Dim reportHeader = <Border VerticalAlignment="Top" HorizontalAlignment="Center" Height="80">
                               <StackPanel>
                                   <TextBlock TextAlignment="Center" FontSize="18" FontWeight="Bold" Foreground="DarkBlue">WPF Fixed Document Reports</TextBlock>
                                   <TextBlock TextAlignment="Center" FontSize="18" FontWeight="Bold" Foreground="DarkBlue"><%= headerText %></TextBlock>
                               </StackPanel>
                           </Border>
        Dim pageHeader = <Border VerticalAlignment="Top" HorizontalAlignment="Stretch">
                             <Grid TextBlock.FontSize="12" Background="LightGray">
                                 <Grid.ColumnDefinitions>
                                     <%= From item In info Select <ColumnDefinition Width=<%= columnWidths(propertyMap(item.Name)) & "*" %>></ColumnDefinition> %>
                                 </Grid.ColumnDefinitions>
                                 <%= From item In info Select
                                     <Border Grid.Column=<%= propertyMap(item.Name) %> BorderBrush="Black" BorderThickness="0.5">
                                         <TextBlock TextAlignment="Center" VerticalAlignment="Center" Grid.Column="0" TextWrapping="Wrap" FontWeight="Bold"><%= columnHeader(propertyMap(item.Name)) %></TextBlock>
                                     </Border> %>
                             </Grid>
                         </Border>

        Dim items = From row In tbl
                    Select <Border VerticalAlignment="Top" HorizontalAlignment="Stretch">
                               <Grid TextBlock.FontSize="12">
                                   <Grid.ColumnDefinitions>
                                       <%= From item In info Select <ColumnDefinition Width=<%= columnWidths(propertyMap(item.Name)) & "*" %>></ColumnDefinition> %>
                                   </Grid.ColumnDefinitions>
                                   <%= From item In info Select
                                       <Border Grid.Column=<%= propertyMap(item.Name) %> BorderBrush="Black" BorderThickness="0.5">
                                           <%= If(TypeOf (CallByName(row, item.Name, CallType.Get)) Is XElement,
                                               CallByName(row, item.Name, CallType.Get),
                                               <TextBlock TextWrapping="Wrap" Margin="5 0" TextAlignment=<%= columnAlignments(propertyMap(item.Name)) %>><%= CallByName(row, item.Name, CallType.Get) %></TextBlock>) %>
                                       </Border> %>
                               </Grid>
                           </Border>
        _currentDoc = New FixedDocumentReport With {.ReportHeader = reportHeader, .PageHeader = pageHeader, .Items = items,
                                                    .PageWidth = 816, .PageHeight = 1056, .PageMargin = New Thickness(48)}

        Return _currentDoc.GenerateReport()
    End Function

    Public Shared Function GridReport(headerText As String, tbl As IEnumerable, gridRows() As Integer, gridColumns() As Integer, columnWidths() As Double, columnHeader() As String, columnAlignments() As TextAlignment) As FixedDocument
        Dim info() As PropertyInfo = (From a In tbl).First.GetType().GetProperties()
        Dim propertyMap As New SortedList(Of String, Integer)
        For Each prop In info
            propertyMap.Add(prop.Name, propertyMap.Count)
        Next

        Dim reportHeader = <Border VerticalAlignment="Top" HorizontalAlignment="Center" Height="80">
                               <StackPanel>
                                   <TextBlock TextAlignment="Center" FontSize="18" FontWeight="Bold" Foreground="DarkBlue">WPF Fixed Document Reports</TextBlock>
                                   <TextBlock TextAlignment="Center" FontSize="18" FontWeight="Bold" Foreground="DarkBlue"><%= headerText %></TextBlock>
                               </StackPanel>
                           </Border>
        Dim pageHeader = <Border VerticalAlignment="Top" HorizontalAlignment="Stretch">
                             <Grid TextBlock.FontSize="12" Background="LightGray">
                                 <Grid.ColumnDefinitions>
                                     <%= From item In info Select <ColumnDefinition Width=<%= columnWidths(propertyMap(item.Name)) & "*" %>></ColumnDefinition> %>
                                 </Grid.ColumnDefinitions>
                                 <%= From item In info Select
                                     <Border Grid.Column=<%= propertyMap(item.Name) %> BorderBrush="Black" BorderThickness="0.5">
                                         <TextBlock TextAlignment="Center" VerticalAlignment="Center" Grid.Column="0" TextWrapping="Wrap" FontWeight="Bold"><%= columnHeader(propertyMap(item.Name)) %></TextBlock>
                                     </Border> %>
                             </Grid>
                         </Border>

        Dim items = From row In tbl
                    Select <Border VerticalAlignment="Top" HorizontalAlignment="Stretch">
                               <Grid TextBlock.FontSize="12">
                                   <Grid.ColumnDefinitions>
                                       <%= From item In info Select <ColumnDefinition Width=<%= columnWidths(propertyMap(item.Name)) & "*" %>></ColumnDefinition> %>
                                   </Grid.ColumnDefinitions>
                                   <%= From item In info Select
                                       <Border Grid.Column=<%= propertyMap(item.Name) %> BorderBrush="Black" BorderThickness="0.5">
                                           <%= If(TypeOf (CallByName(row, item.Name, CallType.Get)) Is XElement,
                                               CallByName(row, item.Name, CallType.Get),
                                               <TextBlock TextWrapping="Wrap" Margin="5 0" TextAlignment=<%= columnAlignments(propertyMap(item.Name)) %>><%= CallByName(row, item.Name, CallType.Get) %></TextBlock>) %>
                                       </Border> %>
                               </Grid>
                           </Border>
        _currentDoc = New FixedDocumentReport With {.ReportHeader = reportHeader, .PageHeader = pageHeader, .Items = items,
                                                    .PageWidth = 816, .PageHeight = 1056, .PageMargin = New Thickness(48)}

        Return _currentDoc.GenerateReport()
    End Function


    Public Shared Function ComplexTableReport(headerText As String, tbl As IEnumerable, columnWidths() As Double, columnHeader() As String, columnAlignments() As TextAlignment) As FixedDocument
        Dim info() As PropertyInfo = (From a In tbl).First.GetType().GetProperties()
        Dim propertyMap As New SortedList(Of String, Integer)
        For Each prop In info
            propertyMap.Add(prop.Name, propertyMap.Count)
        Next

        Dim reportHeader = <Border VerticalAlignment="Top" HorizontalAlignment="Center" Height="80">
                               <StackPanel>
                                   <TextBlock TextAlignment="Center" FontSize="18" FontWeight="Bold" Foreground="DarkBlue">WPF Fixed Document Reports</TextBlock>
                                   <TextBlock TextAlignment="Center" FontSize="18" FontWeight="Bold" Foreground="DarkBlue"><%= headerText %></TextBlock>
                               </StackPanel>
                           </Border>
        Dim pageHeader = <Border VerticalAlignment="Top" HorizontalAlignment="Stretch">
                             <Grid TextBlock.FontSize="12" Background="LightGray">
                                 <Grid.ColumnDefinitions>
                                     <%= From item In info Select <ColumnDefinition Width=<%= columnWidths(propertyMap(item.Name)) & "*" %>></ColumnDefinition> %>
                                 </Grid.ColumnDefinitions>
                                 <%= From item In info Select
                                     <Border Grid.Column=<%= propertyMap(item.Name) %> BorderBrush="Black" BorderThickness="0.5">
                                         <TextBlock TextAlignment="Center" VerticalAlignment="Center" Grid.Column="0" TextWrapping="Wrap" FontWeight="Bold"><%= columnHeader(propertyMap(item.Name)) %></TextBlock>
                                     </Border> %>
                             </Grid>
                         </Border>

        Dim items = From row In tbl
           Select <Border VerticalAlignment="Top" HorizontalAlignment="Stretch">
                      <Grid TextBlock.FontSize="12">
                          <Grid.ColumnDefinitions>
                              <%= From item In info Select <ColumnDefinition Width=<%= columnWidths(propertyMap(item.Name)) & "*" %>></ColumnDefinition> %>
                          </Grid.ColumnDefinitions>
                          <%= From item In info Select
                              <Border Grid.Column=<%= propertyMap(item.Name) %> BorderBrush="Black" BorderThickness="0.5">
                                  <TextBlock Grid.Column="0" TextWrapping="Wrap" Margin="5 0" TextAlignment=<%= columnAlignments(propertyMap(item.Name)) %>><%= CallByName(row, item.Name, CallType.Get) %></TextBlock>
                              </Border> %>
                      </Grid>
                  </Border>
        _currentDoc = New FixedDocumentReport With {.ReportHeader = reportHeader, .PageHeader = pageHeader, .Items = items,
                                                    .PageWidth = 816, .PageHeight = 1056, .PageMargin = New Thickness(48)}

        Return _currentDoc.GenerateReport()
    End Function

    Shared Event ReportProgress(itemCount As Integer, currentItem As Integer)
    Shared Event ReportCompleted()
End Class