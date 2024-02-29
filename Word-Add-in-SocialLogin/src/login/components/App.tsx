import * as React from "react";
import { Button, Field, makeStyles } from "@fluentui/react-components";
import { createClient } from "@supabase/supabase-js";

const useStyles = makeStyles({
  btn_img: {
    width: "1.25rem",
    height: "1.25rem",
    marginRight: "0.5rem",
  },
});

interface AppProps {
  title: string;
}

const supabase = createClient("https://<supabase-url>.supabase.co", "public-anon-key");

const handleSignInWithGoogle = () => {
  supabase.auth.signInWithOAuth({
    provider: "google",
    options: {
      redirectTo: "https://localhost:3000/login.html",
      queryParams: {
        // Ask to select which google account to use every time
        prompt: "consent",
      },
    },
  });
};

const App = (props: AppProps) => {
  const styles = useStyles();
  return (
    <div>
      <Field>{props.title}</Field>
      <div id="google_btn">
        <div className="min-h-screen bg-base-200 flex items-center">
          <div className="card mx-auto w-full max-w-5xl  shadow-xl">
            <div className="grid  md:grid-cols-2 grid-cols-1  bg-base-100 rounded-xl">
              <div className=""></div>
              <div className="py-24 px-10">
                <h2 className="text-2xl font-semibold mb-2 text-center">Login</h2>
                <Button className="btn mt-2 w-full btn-google" onClick={handleSignInWithGoogle}>
                  <img src="../../assets/btn_google_logo.svg" className={styles.btn_img} title="Sign In With Google" />
                  Sign In With Google
                </Button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
