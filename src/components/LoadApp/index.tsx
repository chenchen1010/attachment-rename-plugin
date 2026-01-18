import { ReactElement, useEffect, useState } from "react"
import { bitable } from "@lark-base-open/js-sdk"
import './style.css'

export default function LoadApp(props: { neverShowBanner?: boolean, children: ReactElement }): ReactElement {
  const [loadErr, setLoadErr] = useState(false)

  const TopBanner = <div>
    <div className='errTop'>
      After running the project, please get the webview address and paste it into the Base table "Extended Script" for use. See:&nbsp;
      <a target='_blank' href='https://bytedance.feishu.cn/docx/HazFdSHH9ofRGKx8424cwzLlnZc'>Development Guide</a>
    </div>
  </div>
  useEffect(() => {
    if (props.neverShowBanner) return;
    const timer = new Promise((resolve, reject) => {
      setTimeout(() => {
        reject(false)
      }, 3000)
    })
    const languagePromise = bitable?.bridge?.getLanguage ? bitable.bridge.getLanguage() : Promise.reject(false);
    Promise.race([languagePromise, timer]).then((v) => {
      setLoadErr(false)
    }).catch(() => {
      setLoadErr(true)
    })
  }, [])

  if (props.neverShowBanner) {
    return props.children || null
  }

  return <div>
    {loadErr && TopBanner}
    {props.children}
  </div>
}
