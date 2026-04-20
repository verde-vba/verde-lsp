use std::process::Stdio;
use std::time::Duration;
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};
use tokio::process::{ChildStdin, ChildStdout, Command};
use tokio::time::timeout;

const LSP_TIMEOUT: Duration = Duration::from_secs(5);

fn frame(value: &serde_json::Value) -> Vec<u8> {
    let body = value.to_string();
    let mut out = format!("Content-Length: {}\r\n\r\n", body.len()).into_bytes();
    out.extend_from_slice(body.as_bytes());
    out
}

async fn send(stdin: &mut ChildStdin, value: serde_json::Value) {
    stdin.write_all(&frame(&value)).await.unwrap();
    stdin.flush().await.unwrap();
}

async fn recv(reader: &mut BufReader<ChildStdout>, expected_id: u64) -> serde_json::Value {
    loop {
        let mut content_length = 0usize;
        loop {
            let mut line = String::new();
            reader.read_line(&mut line).await.unwrap();
            let trimmed = line.trim_end_matches(['\r', '\n']);
            if trimmed.is_empty() {
                break;
            }
            if let Some(val) = trimmed.strip_prefix("Content-Length: ") {
                content_length = val.parse().unwrap();
            }
        }
        let mut body = vec![0u8; content_length];
        reader.read_exact(&mut body).await.unwrap();
        let msg: serde_json::Value = serde_json::from_slice(&body).unwrap();
        if msg.get("id").and_then(|v| v.as_u64()) == Some(expected_id) {
            return msg;
        }
        // notifications (no "id") and other responses are skipped
    }
}

#[tokio::test]
async fn stdio_lifecycle_completes_gracefully() {
    let bin = env!("CARGO_BIN_EXE_verde-lsp");
    let mut child = Command::new(bin)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .expect("failed to spawn verde-lsp");

    let mut stdin = child.stdin.take().unwrap();
    let mut reader = BufReader::new(child.stdout.take().unwrap());

    let lifecycle = timeout(LSP_TIMEOUT, async {
        // 1. initialize
        send(
            &mut stdin,
            serde_json::json!({
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "processId": null,
                    "rootUri": null,
                    "capabilities": {}
                }
            }),
        )
        .await;

        let init_resp = recv(&mut reader, 1).await;
        assert!(
            init_resp["result"]["capabilities"].is_object(),
            "initialize must return server capabilities"
        );

        // 2. initialized (notification — no response expected)
        send(
            &mut stdin,
            serde_json::json!({ "jsonrpc": "2.0", "method": "initialized", "params": {} }),
        )
        .await;

        // 3. didOpen a minimal .bas file
        send(
            &mut stdin,
            serde_json::json!({
                "jsonrpc": "2.0",
                "method": "textDocument/didOpen",
                "params": {
                    "textDocument": {
                        "uri": "file:///test.bas",
                        "languageId": "vba",
                        "version": 1,
                        "text": "Sub Hello()\n    Dim x As String\nEnd Sub\n"
                    }
                }
            }),
        )
        .await;

        // 4. completion — position inside Sub body, after "Dim x"
        send(
            &mut stdin,
            serde_json::json!({
                "jsonrpc": "2.0",
                "id": 2,
                "method": "textDocument/completion",
                "params": {
                    "textDocument": { "uri": "file:///test.bas" },
                    "position": { "line": 1, "character": 14 }
                }
            }),
        )
        .await;

        let comp_resp = recv(&mut reader, 2).await;
        let items = comp_resp["result"]["items"]
            .as_array()
            .or_else(|| comp_resp["result"].as_array())
            .expect("completion result must be CompletionList or array");
        assert!(
            !items.is_empty(),
            "expected at least one completion item; got: {comp_resp}"
        );

        // 5. shutdown
        send(
            &mut stdin,
            serde_json::json!({ "jsonrpc": "2.0", "id": 3, "method": "shutdown" }),
        )
        .await;

        let shutdown_resp = recv(&mut reader, 3).await;
        assert!(
            shutdown_resp["result"].is_null(),
            "shutdown must respond with null result"
        );

        // 6. exit notification
        send(
            &mut stdin,
            serde_json::json!({ "jsonrpc": "2.0", "method": "exit" }),
        )
        .await;

        // Close stdin so the server's serve() loop sees EOF and returns
        drop(stdin);
    })
    .await;

    assert!(lifecycle.is_ok(), "LSP lifecycle timed out (> 5 s)");

    let status = timeout(LSP_TIMEOUT, child.wait())
        .await
        .expect("child.wait() timed out")
        .expect("child.wait() failed");

    // Per LSP spec: exit code 0 when shutdown was received before exit
    assert_eq!(
        status.code(),
        Some(0),
        "verde-lsp must exit 0 after shutdown+exit"
    );
}
