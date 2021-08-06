# office-fable
Fable bindings for Office-js

## How was this created

- Install ts2fable (https://github.com/fable-compiler/ts2fable)
  - `npm install -g ts2fable`
- Get office.js typescript file
  - `npm install @types/office-js`
  - node_modules\@types\office-js\index.d.ts
- Create Fable bindings with 
`ts2fable "node_modules\@types\office-js\index.d.ts" src\office-fable\OfficeJS.fs`
- Add Fable.Core to .fsproj `<PackageReference Include="Fable.Core" Version="3.2.8" />`
- Add Fable.React to .fsproj
`<PackageReference Include="Fable.React" Version="7.4.1" />`
- Replace all `PromiseLike` in OfficeJS.fs to `Promise`.
- ts2fable translates union cases which start with a number to start with an underscore. Both is not valid in fsharp. Therefore all union cases starting with an `_` are replaced to start with an `N`. Exmp.: `_3DColumnClustered` -> `N3DColumnClustered` 
- Change union case `[<CompiledName "Tags">] Tags` to `Tags2`.
