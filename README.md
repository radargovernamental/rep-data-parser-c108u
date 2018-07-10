# rep-data-parser-c108u

> Data parser for REP (Relógio de Ponto) data, model C108U. It's a small project for internal use only.

## Install

```bash
yarn install
```

## Usage

You must have a file named `input.xlsx`, containing the first page as the `Registros` page of the extracted file from the device.
After that, just run:

```bash
yarn start
```

## Considerations

We are considering only entries of two or four records. Any entry with 5+ records must be adjusted or will be ignored.


## Contributors

<table>
  <tbody>
    <tr>
      <td align="center">
        <a href="https://github.com/brunocramos">
          <img src="https://avatars.githubusercontent.com/u/4956907?v=4" title="brunocramos" width="80" height="80"><br />
          @brunocramos
        </a>
      </td>
    </tr>
  </tbody>
</table>

## License

MIT © [radargovernamental](https://github.com/radargovernamental/rep-data-parser-c108u)
