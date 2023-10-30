use axum::{Router, routing::get, response::Html, extract::WebSocketUpgrade};
use dioxus::prelude::*;

#[inline_props]
fn ControlledInput<'a>(cx: Scope<'a>, element_id: &'a str, label: &'a str, value: &'a str, on_input: EventHandler<'a, FormEvent>) -> Element {
    render! {
        div {
            label {
                "for": "{element_id}",
                "{label}",
            }
            input {
                id: "{element_id}",
                value: "{value}",
                oninput: move |evt| on_input.call(evt),
            }
        }
    }
}

fn LoginElement(cx: Scope) -> Element {
    let name = use_state(cx, || "".to_string());
    let password = use_state(cx, || "".to_string());
    render! {
        form {
            ControlledInput {
                element_id: "username",
                label: "Baserow username",
                value: name,
                on_input: move |evt: FormEvent| {name.set(evt.value.clone());},
            },
            ControlledInput {
                element_id: "password",
                label: "Baserow password",
                value: password,
                on_input: move |evt: FormEvent| {password.set(evt.value.clone());},
            },
            input { r#type: "submit" },
        }
    }
}

#[tokio::main]
async fn main() {
    let addr: std::net::SocketAddr = ([10, 43, 61, 104], 8881).into();

    let view = dioxus_liveview::LiveViewPool::new();

    let app = Router::new()
        // The root route contains the glue code to connect to the WebSocket
        .route(
            "/",
            get(move || async move {
                Html(format!(
                    r#"
                <!DOCTYPE html>
                <html>
                <head> <title>Dioxus LiveView with Axum</title>  </head>
                <body> <div id="main"></div> </body>
                {glue}
                </html>
                "#,
                    // Create the glue code to connect to the WebSocket on the "/ws" route
                    glue = dioxus_liveview::interpreter_glue(&format!("ws://{addr}/ws"))
                ))
            }),
        )
        // The WebSocket route is what Dioxus uses to communicate with the browser
        .route(
            "/ws",
            get(move |ws: WebSocketUpgrade| async move {
                ws.on_upgrade(move |socket| async move {
                    // When the WebSocket is upgraded, launch the LiveView with the app component
                    _ = view.launch(dioxus_liveview::axum_socket(socket), app).await;
                })
            }),
        );

    println!("Listening on http://{addr}");

    axum::Server::bind(&addr.to_string().parse().unwrap())
        .serve(app.into_make_service())
        .await
        .unwrap();
}

fn app(cx: Scope) -> Element {
    cx.render(rsx! {
        main {
            LoginElement {}
        }
    })
}
