# 必要なアセンブリをロード
Add-Type -Path "C:\path\to\Oracle.ManagedDataAccess.dll"

# Oracleデータベースへの接続情報
$connectionString = "User Id=myUsername;Password=myPassword;Data Source=myDataSource"

# SELECTクエリ
$selectQuery = "SELECT column1, column2 FROM source_table"

# INSERTクエリ
$insertQuery = "INSERT INTO destination_table (column1, column2) VALUES (:value1, :value2)"

try {
    # データベース接続を開く
    $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)
    $connection.Open()

    # SELECTコマンドを作成
    $selectCommand = $connection.CreateCommand()
    $selectCommand.CommandText = $selectQuery

    # データリーダーを使用してデータを1件ずつ取得
    $reader = $selectCommand.ExecuteReader()

    while ($reader.Read()) {
        # 各列の値を取得し、適切な型に変換
        $column1Value = [string]$reader["column1"]
        $column2Value = [int]$reader["column2"]

        # INSERTコマンドを作成
        $insertCommand = $connection.CreateCommand()
        $insertCommand.CommandText = $insertQuery

        # パラメータを設定
        $param1 = New-Object Oracle.ManagedDataAccess.Client.OracleParameter("value1", [Oracle.ManagedDataAccess.Client.OracleDbType]::Varchar2)
        $param1.Value = $column1Value
        $insertCommand.Parameters.Add($param1)

        $param2 = New-Object Oracle.ManagedDataAccess.Client.OracleParameter("value2", [Oracle.ManagedDataAccess.Client.OracleDbType]::Int32)
        $param2.Value = $column2Value
        $insertCommand.Parameters.Add($param2)

        # データを挿入
        $insertCommand.ExecuteNonQuery()
    }

    # データリーダーを閉じる
    $reader.Close()

    # データベース接続を閉じる
    $connection.Close()
}
catch {
    Write-Host "An error occurred: $_"
}
finally {
    if ($connection.State -eq 'Open') {
        $connection.Close()
    }
}
