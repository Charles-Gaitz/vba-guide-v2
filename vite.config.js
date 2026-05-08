import { resolve } from 'path'
export default {
  build: {
    rollupOptions: {
      input: {
        main:                resolve(__dirname, 'index.html'),
        foundations:         resolve(__dirname, 'src/modules/foundations.html'),
        programmingConcepts: resolve(__dirname, 'src/modules/programming-concepts.html'),
        variables:           resolve(__dirname, 'src/modules/variables.html'),
        loops:               resolve(__dirname, 'src/modules/loops.html'),
        references:          resolve(__dirname, 'src/modules/references.html'),
        filters:             resolve(__dirname, 'src/modules/filters.html'),
        debugging:           resolve(__dirname, 'src/modules/debugging.html'),
        calculationsAndDates: resolve(__dirname, 'src/modules/calculations-and-dates.html'),
        pseudocode:          resolve(__dirname, 'src/modules/pseudocode.html'),
        practiceProject:     resolve(__dirname, 'src/modules/practice-project.html'),
        objects:             resolve(__dirname, 'src/modules/objects.html'),
      }
    }
  }
}
